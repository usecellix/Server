import { ForbiddenException, Injectable, Logger, NotFoundException } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { AgentRun, AgentRunDocument, AGENT_RUN_TTL_MS } from './schemas/agent-run.schema';
import { Action, SubTask, WorkbookContext } from './types/agent.types';
import { computeExecutionWaves } from './utils/task-graph.util';

export type WaveDecision = 'accepted' | 'rejected' | 'skipped';

/**
 * Structural check that a readback entry is actually a `SheetContext`
 * (`agent.types.ts`), not the frontend's own `WorkbookContext.sheets` shape
 * (`SheetSnapshot` — `sheetName`/`colCount`/`headers`/`sampleData`, no
 * `.name`/`.values`/`.formulas`/`.numberFormats`). The two are easy to
 * conflate because both are called "WorkbookContext.sheets" in their
 * respective codebases, but they are unrelated types with no field overlap
 * beyond `structure`/`headerRowIndex`. See `applyReadback`'s docblock.
 */
function isValidSheetContext(sheet: unknown): sheet is WorkbookContext['sheets'][number] {
  if (!sheet || typeof sheet !== 'object') return false;
  const s = sheet as Record<string, unknown>;
  return (
    typeof s.name === 'string' &&
    s.name.length > 0 &&
    Array.isArray(s.values) &&
    Array.isArray(s.formulas) &&
    Array.isArray(s.numberFormats)
  );
}

export interface CreateRunInput {
  conversationId: string;
  userId?: string;
  traceId: string;
  prompt: string;
  subtasks: SubTask[];
  context: WorkbookContext;
  promptContext?: string;
  conversationHistory?: { role: 'user' | 'assistant'; content: string }[];
  routerAssumption?: string;
  /** Probed host capabilities, persisted so later waves can read them — TASKS.md #269. */
  excelCapabilities?: { dynamicArrays?: boolean };
}

/**
 * Owns the lifecycle of a stepwise Tier 3 run across the N HTTP requests it
 * spans — STEPWISE_EXECUTION.md SD-1. Deliberately separate from
 * AgenticLoopService, which stays a pure per-wave executor with no knowledge
 * that runs persist at all.
 */
@Injectable()
export class AgentRunStateService {
  private readonly logger = new Logger(AgentRunStateService.name);

  constructor(
    @InjectModel(AgentRun.name)
    private readonly agentRunModel: Model<AgentRunDocument>,
  ) {}

  /**
   * Freezes the plan and its dependency waves. Waves are computed once, here —
   * never recomputed mid-run, so a resumed run cannot silently regroup its own
   * remaining work (SD-2).
   */
  async createRun(input: CreateRunInput): Promise<AgentRunDocument> {
    const waves = computeExecutionWaves(input.subtasks);
    const runId = `run_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;

    return this.agentRunModel.create({
      runId,
      conversationId: input.conversationId,
      userId: input.userId,
      traceId: input.traceId,
      prompt: input.prompt,
      status: 'running',
      waveIndex: -1,
      waveTotal: waves.length,
      subtasks: input.subtasks,
      waves: waves.map((wave) => wave.map((subtask) => subtask.id)),
      subtaskStates: input.subtasks.map((subtask) => ({
        subtaskId: subtask.id,
        actions: [],
        completed: false,
      })),
      context: input.context,
      changeSetIds: [],
      promptContext: input.promptContext,
      conversationHistory: input.conversationHistory ?? [],
      routerAssumption: input.routerAssumption,
      excelCapabilities: input.excelCapabilities,
      expiresAt: new Date(Date.now() + AGENT_RUN_TTL_MS),
    });
  }

  /**
   * Loads a run for `/continue`, refusing one that belongs to somebody else.
   * The ownership check is the whole reason `userId` is on the document: a
   * runId is guessable enough that without it, holding one would be enough to
   * drive another user's build.
   */
  /**
   * The unfinished run for a conversation, if there is one — Phase 8 of
   * LONG_PROMPT_RELIABILITY_PLAN.md (TASKS.md #293).
   *
   * A stepwise build spans several HTTP requests, and each one ends by handing
   * control back to the client. If that client never comes back — the task pane
   * lost its connection, Excel was closed, the machine slept — the run simply
   * sits in `awaiting_decision` until its TTL expires, with its finished waves
   * already applied and no way for anyone to continue it. Observed live: a run
   * whose first wave had completed successfully was stranded with no route
   * back, and the only option was to start the whole build again.
   *
   * Long builds are hit hardest for the obvious reason that they are long.
   */
  async findResumableRun(
    conversationId: string,
    userId?: string,
  ): Promise<AgentRunDocument | null> {
    const run = await this.agentRunModel
      .findOne({ conversationId, status: { $in: ['awaiting_decision', 'running'] } })
      .sort({ _id: -1 });

    if (!run) return null;
    // Same ownership discipline as `loadRunForUser`: a run recorded with an
    // owner may only be resumed by that owner; one without predates auth
    // wiring (or came from the eval-bypass path) and stays resumable.
    if (run.userId && userId && run.userId !== userId) return null;
    // Nothing left to do is not resumable, however the status reads.
    return this.nextExecutableWave(run) ? run : null;
  }

  async loadRunForUser(runId: string, userId?: string): Promise<AgentRunDocument> {
    const run = await this.agentRunModel.findOne({ runId });
    if (!run) {
      throw new NotFoundException('RUN_NOT_FOUND');
    }
    // A run recorded with an owner may only be continued by that owner. A run
    // with no owner predates auth wiring (or came from the eval bypass path) —
    // those stay continuable, matching how `Conversation.userId` treats its own
    // orphans rather than locking them out.
    if (run.userId && run.userId !== userId) {
      throw new ForbiddenException('RUN_NOT_YOURS');
    }
    return run;
  }

  /**
   * Records the client's decision on the wave just emitted, and cascade-skips
   * every subtask transitively depending on one that was rejected or skipped
   * (SD-4). Returns the ids that were cascade-skipped so the caller can report
   * them — a skipped step reported as done is exactly the false-completeness
   * failure CODEBASE_ANALYSIS.md §3.7 keeps re-teaching.
   */
  async applyDecision(
    run: AgentRunDocument,
    decision: WaveDecision,
  ): Promise<{ cascadeSkipped: string[] }> {
    const decidedWave = run.waves[run.waveIndex] ?? [];
    for (const subtaskId of decidedWave) {
      const state = run.subtaskStates.find((entry) => entry.subtaskId === subtaskId);
      if (state) state.decision = decision;
    }

    const cascadeSkipped: string[] =
      decision === 'accepted'
        ? []
        : [...this.collectDownstream(run.subtasks, new Set(decidedWave))].filter(
            (id) => !decidedWave.includes(id),
          );

    for (const subtaskId of cascadeSkipped) {
      const state = run.subtaskStates.find((entry) => entry.subtaskId === subtaskId);
      if (!state || state.decision) continue;
      state.decision = 'skipped';
      state.failedReason =
        state.failedReason ?? 'Skipped because a step it depends on was not accepted';
    }

    if (cascadeSkipped.length > 0) {
      this.logger.log(
        `Run ${run.runId}: wave ${run.waveIndex} ${decision} — cascade-skipped [${cascadeSkipped.join(', ')}]`,
      );
    }

    run.markModified('subtaskStates');
    await run.save();
    return { cascadeSkipped };
  }

  /**
   * The next wave that still has work worth doing, skipping any wave whose
   * every subtask has already been decided (cascade-skipped). Returns null when
   * the run is finished.
   */
  nextExecutableWave(run: AgentRunDocument): { waveIndex: number; subtasks: SubTask[] } | null {
    const byId = new Map(run.subtasks.map((subtask) => [subtask.id, subtask]));

    for (let index = run.waveIndex + 1; index < run.waves.length; index += 1) {
      const pending = run.waves[index].filter((subtaskId) => {
        const state = run.subtaskStates.find((entry) => entry.subtaskId === subtaskId);
        return !state?.decision;
      });
      if (pending.length === 0) continue;

      const subtasks = pending
        .map((subtaskId) => byId.get(subtaskId))
        .filter((subtask): subtask is SubTask => Boolean(subtask));
      if (subtasks.length > 0) {
        return { waveIndex: index, subtasks };
      }
    }

    return null;
  }

  /** Persists a wave's executed results and advances the run's cursor. */
  async recordWaveResult(
    run: AgentRunDocument,
    waveIndex: number,
    results: Array<{
      subtaskId: string;
      actions: Action[];
      completed: boolean;
      verified?: boolean;
      failedReason?: string;
    }>,
    changeSetId?: string,
  ): Promise<void> {
    for (const result of results) {
      const state = run.subtaskStates.find((entry) => entry.subtaskId === result.subtaskId);
      if (!state) continue;
      state.actions = result.actions;
      state.completed = result.completed;
      state.verified = result.verified;
      state.failedReason = result.failedReason;
    }

    run.waveIndex = waveIndex;
    if (changeSetId) run.changeSetIds.push(changeSetId);
    run.status = 'awaiting_decision';
    run.expiresAt = new Date(Date.now() + AGENT_RUN_TTL_MS);
    run.markModified('subtaskStates');
    await run.save();
  }

  async markStatus(
    run: AgentRunDocument,
    status: AgentRun['status'],
  ): Promise<void> {
    run.status = status;
    await run.save();
  }

  /**
   * Merges a client-supplied readback of the sheets an accepted wave touched
   * into the run's base context, so the next wave's Executor plans against the
   * OBSERVED result rather than the shadow workbook's prediction — the
   * Shortcut-parity capability TASKS.md #153 is named for.
   *
   * Additive and defensive: an absent or malformed readback leaves the context
   * exactly as it was, degrading to today's predicted-state behaviour rather
   * than failing the run. This is not a hypothetical: the frontend's own
   * `WorkbookContext.sheets` (`SheetSnapshot[]` — `sheetName`, `colCount`,
   * `headers`, `sampleData`, no `.name`/`.values`/`.formulas` at all) is a
   * completely different shape from this backend's `SheetContext[]`, and there
   * is currently no translator between them. Sending that shape through
   * unvalidated corrupted `run.context.sheets` with a malformed entry, which
   * the next wave's Executor then crashed on reading `.values.length` — a live
   * "Cannot read properties of undefined (reading 'length')" failure this
   * guard exists to make structurally impossible rather than merely unlikely.
   */
  async applyReadback(
    run: AgentRunDocument,
    readback: WorkbookContext['sheets'] | undefined,
  ): Promise<void> {
    if (!Array.isArray(readback) || readback.length === 0) return;

    const validSheets = readback.filter(isValidSheetContext);
    if (validSheets.length !== readback.length) {
      this.logger.warn(
        `Run ${run.runId}: rejected readback — ${readback.length - validSheets.length} of ` +
          `${readback.length} sheet(s) did not match the expected SheetContext shape ` +
          `(name/values/formulas/numberFormats as arrays). Continuing with predicted state.`,
      );
    }
    if (validSheets.length === 0) return;

    const bySheetName = new Map(validSheets.map((sheet) => [sheet.name, sheet]));
    const merged = run.context.sheets.map((sheet) => bySheetName.get(sheet.name) ?? sheet);

    for (const sheet of validSheets) {
      if (!merged.some((entry) => entry.name === sheet.name)) {
        merged.push(sheet);
      }
    }

    run.context = { ...run.context, sheets: merged };
    run.markModified('context');
    await run.save();
  }

  /** Subtasks the run never delivered, for an honest closing summary (SD-4). */
  summarizeSkipped(run: AgentRunDocument): Array<{ subtaskId: string; description: string; reason: string }> {
    const byId = new Map(run.subtasks.map((subtask) => [subtask.id, subtask]));
    return run.subtaskStates
      .filter((state) => state.decision === 'rejected' || state.decision === 'skipped')
      .map((state) => ({
        subtaskId: state.subtaskId,
        description: byId.get(state.subtaskId)?.description ?? state.subtaskId,
        reason: state.failedReason ?? `Step was ${state.decision}`,
      }));
  }

  /**
   * Transitive closure of "depends on one of `seedIds`". Same graph walk as
   * AgenticLoopService.collectDownstreamSubtasks and the frontend's
   * collectCascadeRejectIds — three call sites, one rule, deliberately not
   * shared across the repo boundary but kept behaviourally identical.
   */
  private collectDownstream(subtasks: SubTask[], seedIds: Set<string>): Set<string> {
    const downstream = new Set<string>(seedIds);
    let changed = true;
    while (changed) {
      changed = false;
      for (const subtask of subtasks) {
        if (downstream.has(subtask.id)) continue;
        if (subtask.dependsOn.some((dep) => downstream.has(dep))) {
          downstream.add(subtask.id);
          changed = true;
        }
      }
    }
    return downstream;
  }
}
