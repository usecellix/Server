import { dedupeIdenticalWrites } from './utils/dedupe-writes.util';
import { Injectable, Logger } from '@nestjs/common';
import { ExecutorAgent } from './executor.agent';
import { VerifierAgent } from './verifier.agent';
import {
  Action,
  DroppedAction,
  ExecutorOutput,
  SubTask,
  VerifierIssue,
  VerifierOutput,
  WorkbookContext,
} from './types/agent.types';
import { SseEmitter } from './sse.emitter';
import { FormulaAnalyzer } from '../formula/formula.analyzer';
import { FormulaValidatorService } from '../formula/formula-validator.service';
import { FormulaValidationResult } from '../formula/formula.types';
import { buildShadowWorkbook, shadowToWorkbookContext } from '../virtual/shadowWorkbook';
import { ShadowWorkbook } from '../virtual/shadowWorkbook.types';
import { virtualApply } from '../virtual/virtualApply';
import { ToolBridgeService } from './tool-bridge.service';
import { mergeRangeIntoSheet } from './utils/range-merge.util';
import { computeExecutionWaves } from './utils/task-graph.util';
import { cloneActionsForSheet, findCloneGroups } from './utils/template-replicate.util';
import { buildHeaderTableActions } from './utils/header-table-split.util';
import { AdaptiveConcurrency } from './utils/adaptive-concurrency.util';
import { classifyLlmError } from './utils/transient-llm-error.util';
import { CompletenessChecker } from './checkers/completeness.checker';
import { FormattingChecker } from './checkers/formatting.checker';
import { SemanticFormulaChecker } from './checkers/semantic-formula.checker';
import { OverwriteOccupancyChecker } from './checkers/overwrite-occupancy.checker';
import { StructuralIntentChecker } from './checkers/structural-intent.checker';
import { ComputedColumnChecker } from './checkers/computed-column.checker';
import { SpecConformanceChecker } from './checkers/spec-conformance.checker';
import { CheckerResult, mergeCheckerResults } from './checkers/checker.types';
import { buildDeterministicSubtaskActions } from './utils/compound-action.util';
import { StepRetryExhaustedError } from './errors';
import { StepRetryContext } from './types/verifier.types';
import { StructuredLogger } from './logging/structured-logger';
import { shouldSkipVerifier } from './verifier-skip.policy';
import { isDestructiveActionType } from './verifier-skip.policy';
import { isExecutorBlockedSignal } from './utils/verifier-partial-parse.util';
import { rebindFormatRangeNumberFormats } from './utils/preserve-number-format.util';
import { UsageTotals } from './utils/usage-accumulator.util';

export interface AgenticLoopOptions {
  conversationId?: string;
  correlationId?: string;
  toolEmit?: (event: string, data: Record<string, unknown>) => void;
  parseFailureTracker?: { hadFailure: boolean };
  /** Out-param — every Executor/Verifier call this run makes accumulates its
   * real usage here, same object the caller (OrchestratorService) passed to
   * PlannerAgent.plan(), so one run produces one combined total. */
  usageTotals?: UsageTotals;
  /**
   * Fires when the client disconnects/aborts (e.g. "Stop" in the taskpane).
   * Checked between execution waves, same spot as the TIMEOUT_MS check —
   * without it, a cancelled run kept executing every remaining wave (LLM
   * calls included) against a response nobody was reading anymore.
   */
  abortSignal?: AbortSignal;
  /**
   * Called as each execution wave finishes, with the actions THAT WAVE
   * produced — TASKS.md #174 (progressive emission, the "B" half of #153).
   *
   * Lets the caller turn a finished wave into an Accept card immediately
   * instead of holding everything until the whole loop returns. The loop stays
   * ignorant of ChangeSets and SSE cards; it only reports "this much is done".
   *
   * Deliberately awaited: the caller creates a ChangeSet, and letting the next
   * wave start before that resolves would let two waves race to emit cards out
   * of order. A wave is 30-50s of LLM time, so a few ms of ChangeSet work costs
   * nothing measurable.
   *
   * Never allowed to break the run — a throw here is logged and swallowed,
   * because a presentation concern must not destroy work the loop has already
   * done (the TASKS.md #173 lesson).
   */
  onWaveComplete?: (waveActions: Action[], waveIndex: number) => Promise<void>;
  /**
   * Set by `runWave` (STEPWISE_EXECUTION.md) — suppresses the hard `ERROR` SSE
   * event this loop otherwise sends when it times out with nothing to show.
   *
   * That event exists for the one-shot path, where "nothing delivered" really
   * is the end of the request. For a stepwise wave it is not: `executeStepwiseWave`
   * treats a wave that timed out with zero actions as recoverable — it marks
   * the wave skipped and moves on to the next one, which can still succeed.
   *
   * Without this flag, a live run showed the actual failure mode: the loop's
   * raw `ERROR` reached the frontend, whose SSE handler treats `error` as
   * terminal (`runtime.aborted = true`, spinner stops, turn shown as failed) —
   * while the backend kept working for several more minutes underneath,
   * trying the next wave. The user saw "Agentic loop timeout" and gave up on
   * a turn that was, from the backend's perspective, still in progress and
   * might have gone on to succeed. `VERIFY_FAIL`/`THINKING` events (both
   * already non-fatal per `SseEmitter`) still fire, so the wave's own status
   * line is not silent — only the hard-stop signal is withheld.
   */
  isStepwiseWave?: boolean;
}

export interface CompletedSubtaskResult {
  subtaskId: string;
  actions: Action[];
  verified: boolean;
}

export interface FailedSubtaskResult {
  subtaskId: string;
  reason: string;
}

export interface AgenticLoopResult {
  actions: Action[];
  iterationsRun: number;
  verifierPassed: boolean;
  completedSubtasks: CompletedSubtaskResult[];
  /**
   * The single MOST RELEVANT failure — kept for callers/messages that only
   * ever describe one ("a later step could not be completed: ..."). This is
   * `failedSubtasks[0]` when non-empty, `null` otherwise; it is NOT
   * necessarily the only failure. TASKS.md #195 found that a wave of many
   * independent parallel subtasks (e.g. 12 month-sheet creates with no
   * dependsOn between them) can have SEVERAL genuinely fail at once — using
   * only this field silently discarded every failure but one, with no
   * recorded reason anywhere for the rest. Prefer `failedSubtasks` for
   * anything that needs to know the true failure count or list every reason.
   */
  failedSubtask: FailedSubtaskResult | null;
  /** Every subtask that failed this run, not just the first — see `failedSubtask`'s docblock. */
  failedSubtasks: FailedSubtaskResult[];
  /** True when some subtasks completed but the full chain did not verify/pass. */
  partialProgress: boolean;
}

interface SubtaskActionState {
  subtask: SubTask;
  actions: Action[];
  /** Actions the Executor emitted that normalization rejected — verified against, never ignored. */
  droppedActions: DroppedAction[];
  completed: boolean;
  verified?: boolean;
  failedReason?: string;
}

type RetryContext = Pick<
  WorkbookContext,
  'verifierFeedback' | 'verifierIssues' | 'formulaValidationFeedback' | 'formulaValidationIssues'
>;

@Injectable()
export class AgenticLoopService {
  private readonly logger = new Logger(AgenticLoopService.name);
  private readonly MAX_ITERATIONS_PER_SUBTASK = 10;
  private readonly MAX_STEP_RETRIES = 2;
  private readonly MAX_FORMULA_RETRIES = 2;
  private readonly MAX_TOOL_REQUESTS = 5;
  /**
   * Wall-clock budget for the whole execute+verify loop, measured from loop
   * start (the Planner's own time is NOT counted against it).
   *
   * Raised from 300_000 on evidence, not preference: a live 20-subtask ledger
   * build ran 542s end to end and died here with nothing to show. Its Main-sheet
   * subtasks form a ~6-deep dependency chain (create -> headers -> Jan-Jun
   * formulas -> Jul-Dec -> KPI row -> consolidated header/chart), and each level
   * is a serial LLM round trip that no amount of sibling parallelism can
   * shorten. Six levels at 30-50s each already approaches 300s before the
   * verifier runs at all.
   *
   * This is a mitigation, not a fix. The real answer is TASKS.md #153's
   * resumable loop, where a long build stops being one connection holding one
   * budget. Until then a build that would have finished at 320s must not be
   * thrown away at 300s.
   */
  private readonly TIMEOUT_MS = 480_000;
  /**
   * Max Executor subtasks run concurrently within one wave. TASKS.md #275 —
   * live evidence from two consecutive real runs of the same 12-month-sheet
   * prompt: EVERY one of the 12 independent month-create subtasks (grouped
   * into one parallel wave by `computeExecutionWaves`, since they have no
   * dependsOn between them) failed to reach `completed: true` in both runs,
   * with a DIFFERENT subset naming `OpenRouter could not verify available
   * credits for this request in time` and the rest exhausting
   * MAX_ITERATIONS_PER_SUBTASK with no distinct reason — the signature of 12
   * simultaneous LLM calls fighting over the provider's concurrency ceiling
   * rather than 12 independent failures. Firing all of them via one
   * unthrottled `Promise.all` (previously no cap existed) is what created the
   * contention. Capping how many run at once turns "12 at once, most fail"
   * into "a few at once, queued", trading wall-clock time for actually
   * finishing — the workbook this run built had 4-6 real month sheets missing
   * entirely because their subtask never even produced one action.
   */
  /** Starting width for a wave; `AdaptiveConcurrency` moves it from here. */
  private readonly MAX_WAVE_CONCURRENCY = 3;

  constructor(
    private readonly executor: ExecutorAgent,
    private readonly verifier: VerifierAgent,
    private readonly formulaAnalyzer: FormulaAnalyzer,
    private readonly formulaValidator: FormulaValidatorService,
    private readonly toolBridge: ToolBridgeService,
    private readonly completenessChecker: CompletenessChecker,
    private readonly formattingChecker: FormattingChecker,
    private readonly overwriteOccupancyChecker: OverwriteOccupancyChecker,
    private readonly semanticFormulaChecker: SemanticFormulaChecker = new SemanticFormulaChecker(),
    private readonly structuredLogger: StructuredLogger = new StructuredLogger(),
    private readonly structuralIntentChecker: StructuralIntentChecker = new StructuralIntentChecker(),
    private readonly computedColumnChecker: ComputedColumnChecker = new ComputedColumnChecker(),
    private readonly specConformanceChecker: SpecConformanceChecker = new SpecConformanceChecker(),
  ) {}

  async run(
    originalPrompt: string,
    subtasks: SubTask[],
    context: WorkbookContext,
    emitter: SseEmitter,
    loopOptions: AgenticLoopOptions = {},
  ): Promise<AgenticLoopResult> {
    return this.runInternal(originalPrompt, subtasks, context, emitter, loopOptions);
  }

  /**
   * Executes and verifies ONE dependency wave, for stepwise Tier 3 runs
   * (STEPWISE_EXECUTION.md SD-1/SD-3). The whole loop — Executor iterations,
   * shadow-workbook dry run, deterministic checkers, scoped retry — runs
   * exactly as it does for a full run; the only difference is the set of
   * subtasks it is handed and the already-decided work it is told about.
   *
   * `priorActions` are the actions earlier accepted waves already produced.
   * They are seeded as completed, invisible-to-retry state so this wave's
   * Executor and the shadow workbook both see the sheets those waves created —
   * without them, "populate Main" would plan against a workbook where Main does
   * not exist.
   *
   * Deliberately delegates to `runInternal` rather than reimplementing the
   * sequencing: a second copy of the verify/retry logic is exactly the
   * two-implementations-of-one-rule drift this codebase keeps finding bugs in.
   */
  async runWave(
    originalPrompt: string,
    waveSubtasks: SubTask[],
    priorActions: Array<{ subtask: SubTask; actions: Action[] }>,
    context: WorkbookContext,
    emitter: SseEmitter,
    loopOptions: AgenticLoopOptions = {},
  ): Promise<AgenticLoopResult> {
    return this.runInternal(
      originalPrompt,
      waveSubtasks,
      context,
      emitter,
      { ...loopOptions, isStepwiseWave: true },
      priorActions,
    );
  }

  private async runInternal(
    originalPrompt: string,
    subtasks: SubTask[],
    context: WorkbookContext,
    emitter: SseEmitter,
    loopOptions: AgenticLoopOptions,
    /**
     * Stepwise runs only (STEPWISE_EXECUTION.md SD-1) — work earlier accepted
     * waves already produced. Seeded as completed state so the shadow workbook
     * and Executor context see those sheets/values, then filtered back out of
     * every result, so this wave is verified and reported on its own terms.
     */
    priorActions: Array<{ subtask: SubTask; actions: Action[] }> = [],
  ): Promise<AgenticLoopResult> {
    const startedAt = Date.now();
    let iterationsRun = 0;
    let timedOut = false;
    let cancelled = false;
    const formulaValidationLog: FormulaValidationResult[] = [];

    const ordered = this.orderByDependencies(subtasks);
    const priorIds = new Set(priorActions.map((entry) => entry.subtask.id));
    const priorStates: SubtaskActionState[] = priorActions.map((entry) => ({
      subtask: entry.subtask,
      actions: entry.actions,
      droppedActions: [],
      completed: true,
      verified: true,
    }));
    const ownStates: SubtaskActionState[] = ordered.map((subtask) => ({
      subtask,
      actions: [],
      droppedActions: [],
      completed: false,
    }));
    // Prior states come FIRST so shadow-workbook replay applies them before
    // this wave's own actions — the dependency order they were accepted in.
    const subtaskStates: SubtaskActionState[] = [...priorStates, ...ownStates];

    const waves = computeExecutionWaves(ordered);
    // Prior waves' subtasks are already done: naming them completed is what
    // makes their actions visible to this wave's Executor context.
    const completedIds = new Set<string>(priorIds);
    // Phase 3 (TASKS.md #286) — one controller per run, so a provider that
    // pushed back on an earlier wave is still being treated gently on the next
    // one rather than starting wide again every time.
    const concurrency = new AdaptiveConcurrency({
      start: this.MAX_WAVE_CONCURRENCY,
      floor: 1,
      ceiling: 6,
    });

    for (const wave of waves) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        timedOut = true;
        break;
      }

      if (loopOptions.abortSignal?.aborted) {
        cancelled = true;
        timedOut = true;
        this.logger.warn(
          `Agentic loop cancelled (client disconnected/stopped) before wave — ${completedIds.size} subtask(s) already completed`,
        );
        break;
      }

      if (wave.length > 1) {
        emitter.send({
          type: 'THINKING',
          message: `Running ${wave.length} independent steps in parallel...`,
        });
      }

      // Await every sibling to settlement so a mid-wave LLM failure cannot leave
      // other in-flight fetches as unhandled rejections (Node process crash).
      // Throttled (TASKS.md #275) — a wide wave (e.g. 12 independent
      // month-sheet subtasks) must not fire all its Executor calls at once;
      // that's what was tripping the provider's own concurrency ceiling. The
      // width now adapts to what the provider actually tolerates rather than
      // sitting at a flat guess (Phase 3 / TASKS.md #286).
      const runOne = async (subtask: SubTask) => {
          const state = subtaskStates.find((entry) => entry.subtask.id === subtask.id);
          if (!state) {
            return { iterations: 0 as number, error: null as unknown };
          }

          // A dependency that never actually completed (failed, blocked, or
          // skipped) must not let this subtask run as if it existed — this is
          // what previously let "write formulas on Main" execute against
          // month sheets whose creation step had failed, producing #REF/
          // #VALUE! errors instead of being blocked. TASKS.md #261.
          const unmetDeps = subtask.dependsOn.filter((depId) => !completedIds.has(depId));
          if (unmetDeps.length > 0) {
            state.failedReason = `Skipped — prerequisite step(s) did not complete: ${unmetDeps.join(', ')}`;
            this.logger.warn(
              `Skipping subtask "${subtask.description}" — unmet dependencies: ${unmetDeps.join(', ')}`,
            );
            return { iterations: 0 as number, error: null as unknown };
          }

          // LONG_PROMPT_RELIABILITY_PLAN.md — a deterministic header/table
          // step (split out by `splitSpecPinnedSubtasks`) is built entirely
          // from `expectedHeaders` by code, never the Executor: it cannot
          // drift, time out, or hit the iteration cap, and costs zero LLM
          // calls. TASKS.md #280.
          if (subtask.isDeterministicHeaderStep) {
            state.actions = buildHeaderTableActions(subtask);
            state.droppedActions = [];
            state.completed = true;
            return { iterations: 0 as number, error: null as unknown };
          }
          // A step whose actions were built by code at plan time (the ledger
          // dashboard, TASKS.md #327) — applied verbatim, no model call.
          if (subtask.deterministicActions?.length) {
            state.actions = subtask.deterministicActions.map((action) => ({ ...action }));
            state.droppedActions = [];
            state.completed = true;
            return { iterations: 0 as number, error: null as unknown };
          }

          const visibleIds = new Set([...completedIds, ...subtask.dependsOn]);
          try {
            const iterations = await this.executeSubtask(
              state,
              subtaskStates,
              context,
              emitter,
              startedAt,
              formulaValidationLog,
              loopOptions,
              visibleIds,
              () => {
                timedOut = true;
              },
              originalPrompt,
            );
            return { iterations, error: null as unknown };
          } catch (error) {
            const reason =
              error instanceof Error ? error.message : 'Parallel subtask execution failed';
            state.failedReason = reason;
            this.logger.warn(
              `Parallel subtask "${subtask.description}" failed: ${reason}`,
            );
            return { iterations: 0, error };
          }
      };

      // LONG_PROMPT_RELIABILITY_PLAN.md Phase 2 — siblings that are the same
      // work on different sheets (12 months) run ONE Executor build; the rest
      // are stamped from its accepted actions. A failed template falls back to
      // building each sibling normally, so this can only ever remove work.
      const cloneGroups = findCloneGroups(wave);
      const cloneIds = new Set(cloneGroups.flatMap((group) => group.clones.map((c) => c.id)));
      if (cloneGroups.length > 0) {
        const reused = cloneGroups.reduce((sum, group) => sum + group.clones.length, 0);
        emitter.send({
          type: 'THINKING',
          message: `Building one sheet, then reusing it for ${reused} more...`,
        });
      }

      const waveSettled = await this.runWithAdaptiveConcurrency(
        wave.filter((subtask) => !cloneIds.has(subtask.id)),
        concurrency,
        runOne,
        (result) => (result.error ? classifyLlmError(result.error) : null),
      );

      const fallbackClones: SubTask[] = [];
      for (const group of cloneGroups) {
        const templateState = subtaskStates.find(
          (entry) => entry.subtask.id === group.template.id,
        );
        if (!templateState || !templateState.completed || templateState.failedReason) {
          fallbackClones.push(...group.clones);
          continue;
        }
        for (const clone of group.clones) {
          const cloneState = subtaskStates.find((entry) => entry.subtask.id === clone.id);
          if (!cloneState) continue;
          cloneState.actions = cloneActionsForSheet(
            templateState.actions,
            group.template.targetSheet,
            clone.targetSheet,
          );
          cloneState.droppedActions = [];
          cloneState.completed = true;
        }
        this.logger.log(
          `Reused "${group.template.targetSheet}" build for ${group.clones.length} sibling sheet(s) — ${group.clones.length} Executor call(s) saved`,
        );
      }
      if (fallbackClones.length > 0) {
        this.logger.warn(
          `Template build did not complete — building ${fallbackClones.length} sibling(s) individually`,
        );
        waveSettled.push(
          ...(await this.runWithAdaptiveConcurrency(fallbackClones, concurrency, runOne, (result) =>
            result.error ? classifyLlmError(result.error) : null,
          )),
        );
      }

      iterationsRun += waveSettled.reduce((sum, entry) => sum + entry.iterations, 0);
      for (const subtask of wave) {
        const state = subtaskStates.find((entry) => entry.subtask.id === subtask.id);
        // Only propagate a subtask as "completed" to later waves' visibility
        // when it actually succeeded — a failed/skipped/blocked step must not
        // make dependents believe its sheet/content exists.
        if (state && state.completed && !state.failedReason) {
          completedIds.add(subtask.id);
        }
      }

      // TASKS.md #174 — hand this wave's actions to the caller so a card can be
      // rendered now. Runs BEFORE the progress status below so the card and its
      // "N steps ready" line arrive in a sensible order.
      if (loopOptions.onWaveComplete) {
        // TASKS.md #278 — same bug #274 fixed in buildLoopResult, found in a
        // second place that fix never touched: a live run had January's
        // template time out, and while the FINAL result correctly excluded
        // its actions, THIS progressive per-wave path flattened every
        // subtask's actions with no completion check at all — so May/July's
        // partial in-progress header cells (a subtask that never reached
        // `completed: true`) were shown as an "Applied" Accept card mid-build.
        // Only a subtask's own recorded completion is ground truth here, same
        // principle as #274.
        const waveActions = wave
          .filter((subtask) => {
            const state = subtaskStates.find((entry) => entry.subtask.id === subtask.id);
            return state?.completed === true;
          })
          .flatMap(
            (subtask) =>
              subtaskStates.find((state) => state.subtask.id === subtask.id)?.actions ?? [],
          );
        if (waveActions.length > 0) {
          try {
            await loopOptions.onWaveComplete(waveActions, waves.indexOf(wave));
          } catch (error) {
            this.logger.warn(
              `onWaveComplete failed (continuing — progressive emission must never cost real work): ${
                error instanceof Error ? error.message : String(error)
              }`,
            );
          }
        }
      }

      // Live progress so the UI is not idle while remaining waves run. Counts
      // THIS run's own states only — a stepwise wave must not report earlier
      // accepted waves' work as if it just produced it.
      const readyActions = ownStates.reduce((sum, s) => sum + (s.actions?.length ?? 0), 0);
      const doneSteps = ownStates.filter((s) => s.completed && !s.failedReason).length;
      const failedSteps = ownStates.filter((s) => Boolean(s.failedReason)).length;
      if (readyActions > 0 || doneSteps > 0) {
        emitter.send({
          type: 'CHECKPOINT',
          step:
            failedSteps > 0
              ? `Progress: ${doneSteps} step(s) ready (${readyActions} changes), ${failedSteps} blocked — continuing…`
              : `Progress: ${doneSteps} step(s) ready · ${readyActions} change(s) prepared — continuing…`,
        });
      }

      // Do not rethrow — keep going so completed siblings can surface as partialProgress.
      // Siblings were awaited above so mid-wave LLM aborts cannot become unhandled rejections.
      const failedInWave = waveSettled.filter((entry) => entry.error);
      if (failedInWave.length > 0) {
        emitter.send({
          type: 'THINKING',
          message: `${failedInWave.length} parallel step(s) failed — continuing with remaining work`,
        });
      }
    }

    if (timedOut) {
      // A timeout with work in hand is NOT a fatal error — TASKS.md #173.
      //
      // This unconditionally emitted ERROR, and the frontend's error branch
      // sets `aborted = true` on the turn. So the partial-progress card that
      // `conversation.service.ts` emits moments later was discarded before it
      // could render. Three live runs ended that way: 11 minutes of work, 14
      // subtasks completed, and the user shown nothing but "Agentic loop
      // timeout".
      //
      // Reserve ERROR for the case where there is genuinely nothing to show.
      // Otherwise say what happened as a status and let the partial card
      // through — that is the §3.7 rule pointing the other way for once:
      // reporting honestly here means NOT overstating a partial success as a
      // total failure.
      const deliverable = ownStates.filter(
        (state) => (state.actions?.length ?? 0) > 0,
      ).length;

      this.logger.warn(
        cancelled
          ? `Agentic loop stopped (client disconnected) after ${deliverable} subtask(s) with actions`
          : `Agentic loop timed out before completion (${deliverable} subtask(s) have actions to deliver)`,
      );

      // Cancelled means the client is already gone — writing SSE events to a
      // closed response would throw ("write after end") for no benefit, since
      // nothing is left to read them.
      if (!cancelled) {
        if (deliverable > 0) {
          emitter.send({
            type: 'THINKING',
            message:
              `Ran out of time before finishing every step — ${deliverable} step(s) are ready to review.`,
          });
        } else if (loopOptions.isStepwiseWave) {
          // See AgenticLoopOptions.isStepwiseWave — this wave's own caller
          // treats zero-actions-plus-timeout as recoverable (skip, try the next
          // wave), so the hard terminal signal must not go out here.
          emitter.send({
            type: 'THINKING',
            message: 'This step ran out of time before producing changes — moving to the next one.',
          });
        } else {
          emitter.send({ type: 'ERROR', message: 'Agentic loop timeout' });
        }
      }

      return this.buildLoopResult(ownStates, iterationsRun, false, {
        preferCompletedOnly: true,
        defaultFailReason: cancelled
          ? 'Cancelled — client disconnected before this step finished'
          : 'Agentic loop timed out before completion',
      });
    }

    let verifierPassed = false;
    let verifierCycle = 0;
    const stepRetryAttempts = new Map<string, number>();
    const maxVerifierCycles = Math.max(1, ordered.length * (this.MAX_STEP_RETRIES + 1));
    let retryExhaustedMessage: string | null = null;
    const validatorSummary = this.formulaValidator.summarizeForVerifier(formulaValidationLog);
    loopOptions.parseFailureTracker ??= { hadFailure: false };
    let lastSubtaskVerifyResults: Array<{
      subtaskId: string;
      passed: boolean;
      feedback: string;
      inconclusive?: boolean;
    }> = [];
    /** Spec 17 Bug B: once a subtask passes verification, never re-execute it. */
    const lockedPassIds = new Set<string>();

    while (!verifierPassed && verifierCycle < maxVerifierCycles && !timedOut) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        timedOut = true;
        // See AgenticLoopOptions.isStepwiseWave: the caller of a stepwise wave
        // treats this as recoverable, not terminal — the VERIFY_FAIL emitted
        // just below the loop (non-fatal, per SseEmitter) still tells the user
        // something went wrong with this step.
        if (!loopOptions.isStepwiseWave) {
          emitter.send({ type: 'ERROR', message: 'Agentic loop timeout' });
        }
        break;
      }

      if (loopOptions.abortSignal?.aborted) {
        cancelled = true;
        timedOut = true;
        this.logger.warn(
          'Agentic loop cancelled (client disconnected/stopped) during verification/retry',
        );
        break;
      }

      verifierCycle += 1;
      emitter.send({ type: 'THINKING', message: 'Running deterministic checks...' });

      // Shadow spans EVERY state (prior waves included) so checks see the real
      // accumulated workbook; the checks themselves grade only this run's own
      // subtasks, so a stepwise wave never re-judges work already accepted.
      const shadow = this.buildShadowFromStates(context, subtaskStates);
      const verifyContext = this.enrichContextFromShadow(shadow);
      // The workbook WITHOUT this wave’s own writes. TASKS.md #302: the
      // occupancy checker asks "was this cell already occupied?", and
      // grading it against `verifyContext` means grading a subtask’s
      // actions against a workbook that already contains those very
      // actions — so a step is refused for having already done its work,
      // and every retry is refused the same way. A live run lost all twelve
      // month formula steps to this, the refusal quoting the step’s own
      // formula back at it as the "existing value".
      const priorStates = subtaskStates.filter((entry) => !ownStates.includes(entry));
      const contextBeforeOwnWrites = this.enrichContextFromShadow(
        this.buildShadowFromStates(context, priorStates),
      );
      const cheapChecks = this.runDeterministicChecks(
        originalPrompt,
        ordered,
        ownStates,
        verifyContext,
        context,
        this.sheetsCreatedBy(priorStates),
        contextBeforeOwnWrites,
      );

      if (cheapChecks.passed && !cheapChecks.requiresLlmVerification) {
        verifierPassed = true;
        lastSubtaskVerifyResults = ordered.map((subtask) => ({
          subtaskId: subtask.id,
          passed: true,
          feedback: 'Deterministic checks passed',
        }));
        for (const subtask of ordered) lockedPassIds.add(subtask.id);
        emitter.send({ type: 'VERIFY_PASS' });
        this.logger.log('Skipped LLM verifier — deterministic checks passed cleanly');
        break;
      }

      if (!cheapChecks.passed) {
        this.logger.warn(
          `Deterministic checks failed cycle ${verifierCycle}: ${cheapChecks.feedback}`,
        );
        emitter.send({
          type: 'THINKING',
          message: `Fixing issues: ${cheapChecks.feedback}`,
        });

        lastSubtaskVerifyResults = cheapChecks.subtaskResults.map((result) => ({
          subtaskId: result.subtaskId,
          passed: result.passed,
          feedback: result.feedback,
        }));
        this.lockPassedSubtasks(lastSubtaskVerifyResults, lockedPassIds);

        const failingIds = this.collectFailingIdsForRetry(
          cheapChecks.subtaskResults,
          lockedPassIds,
        );
        if (failingIds.size === 0) {
          break;
        }

        const exhaustedIds = this.findExhaustedSubtasks(failingIds, stepRetryAttempts);
        if (exhaustedIds.size > 0) {
          retryExhaustedMessage = this.buildRetryExhaustedMessage(exhaustedIds, cheapChecks.feedback);
          for (const id of exhaustedIds) {
            const state = subtaskStates.find((entry) => entry.subtask.id === id);
            if (state) {
              state.failedReason =
                state.failedReason ??
                `Could not complete after ${this.MAX_STEP_RETRIES} attempts: ${cheapChecks.feedback}`;
            }
          }
          break;
        }

        const toRetry = ordered.filter((subtask) => failingIds.has(subtask.id));
        this.logger.log(
          `Selective retry (deterministic): re-executing [${[...failingIds].join(', ')}] — locked passes: [${[...lockedPassIds].join(', ')}]`,
        );

        iterationsRun += await this.retrySubtasks(
          toRetry,
          subtaskStates,
          context,
          emitter,
          startedAt,
          formulaValidationLog,
          loopOptions,
          failingIds,
          cheapChecks,
          () => {
            timedOut = true;
          },
          originalPrompt,
          undefined,
          stepRetryAttempts,
        );
        continue;
      }

      emitter.send({ type: 'THINKING', message: 'Verifying semantic correctness...' });

      // This wave's own actions decide whether the LLM verifier is worth
      // running — prior waves' work was already verified when it was accepted.
      const allActions = this.flattenActions(ownStates);
      const hasFormulaActions = allActions.some(
        (action) => action.type === 'SET_FORMULA' || action.type === 'FILL_DOWN',
      );
      const skipDecision = shouldSkipVerifier({
        actions: allActions,
        subtaskCount: ordered.length,
        executorParsedOnFirstAttempt: !loopOptions.parseFailureTracker?.hadFailure,
        hasFormulaActions,
      });

      let verification: VerifierOutput;
      if (skipDecision.skip) {
        this.logger.log(
          `[${loopOptions.correlationId ?? '-'}] ${skipDecision.reason}`,
        );
        verification = {
          passed: true,
          feedback: skipDecision.reason,
          issues: [],
          subtaskResults: ordered.map((subtask) => ({
            subtaskId: subtask.id,
            passed: true,
            feedback: skipDecision.reason,
            issues: [],
          })),
        };
      } else {
        this.logger.log(
          `[${loopOptions.correlationId ?? '-'}] Running Verifier: ${skipDecision.reason}`,
        );
        // Spec 17: only verify subtasks that are not already locked as passed.
        const toVerify = ordered.filter((subtask) => !lockedPassIds.has(subtask.id));
        if (toVerify.length === 0) {
          verification = {
            passed: true,
            feedback: 'All subtasks already verified',
            issues: [],
            subtaskResults: ordered.map((subtask) => ({
              subtaskId: subtask.id,
              passed: true,
              feedback: 'Previously verified',
              issues: [],
            })),
          };
        } else {
          const partial = await this.verifier.verify(
            originalPrompt,
            toVerify,
            this.actionsBySubtaskMap(subtaskStates),
            verifyContext,
            validatorSummary,
            loopOptions.correlationId,
            loopOptions.usageTotals,
          );
          verification = this.mergeWithLockedPasses(partial, ordered, lockedPassIds, lastSubtaskVerifyResults);
        }
      }

      if (verification.passed) {
        verifierPassed = true;
        lastSubtaskVerifyResults = verification.subtaskResults.map((result) => ({
          subtaskId: result.subtaskId,
          passed: result.passed,
          feedback: result.feedback,
          inconclusive: result.inconclusive,
        }));
        this.lockPassedSubtasks(lastSubtaskVerifyResults, lockedPassIds);
        emitter.send({ type: 'VERIFY_PASS' });
        break;
      }

      this.logger.warn(`Verifier failed cycle ${verifierCycle}: ${verification.feedback}`);
      emitter.send({
        type: 'THINKING',
        message: `Fixing issues: ${verification.feedback}`,
      });

      lastSubtaskVerifyResults = verification.subtaskResults.map((result) => ({
        subtaskId: result.subtaskId,
        passed: result.passed,
        feedback: result.feedback,
        inconclusive: result.inconclusive,
      }));
      this.lockPassedSubtasks(lastSubtaskVerifyResults, lockedPassIds);

      const failingIds = this.collectFailingIdsForRetry(
        verification.subtaskResults,
        lockedPassIds,
      );

      // Inconclusive-only: re-verify without re-executing (next loop cycle).
      const inconclusiveOnly =
        failingIds.size === 0 &&
        verification.subtaskResults.some((r) => r.inconclusive && !lockedPassIds.has(r.subtaskId));
      if (inconclusiveOnly) {
        this.logger.warn(
          'Verifier had inconclusive (truncated) results — re-verifying without re-execution',
        );
        continue;
      }

      if (failingIds.size === 0) {
        break;
      }

      const exhaustedIds = this.findExhaustedSubtasks(failingIds, stepRetryAttempts);
      if (exhaustedIds.size > 0) {
        retryExhaustedMessage = this.buildRetryExhaustedMessage(exhaustedIds, verification.feedback);
        for (const id of exhaustedIds) {
          const state = subtaskStates.find((entry) => entry.subtask.id === id);
          if (state) {
            state.failedReason =
              state.failedReason ??
              `Could not complete after ${this.MAX_STEP_RETRIES} attempts: ${verification.feedback}`;
          }
        }
        break;
      }

      const toRetry = ordered.filter((subtask) => failingIds.has(subtask.id));
      this.logger.log(
        `Selective retry: re-executing [${[...failingIds].join(', ')}] — locked passes: [${[...lockedPassIds].join(', ')}]`,
      );

      iterationsRun += await this.retrySubtasks(
        toRetry,
        subtaskStates,
        context,
        emitter,
        startedAt,
        formulaValidationLog,
        loopOptions,
        failingIds,
        {
          passed: false,
          requiresLlmVerification: true,
          feedback: verification.feedback,
          issues: verification.issues,
          subtaskResults: verification.subtaskResults,
        },
        () => {
          timedOut = true;
        },
        originalPrompt,
        (subtask) => {
          const subtaskResult = verification.subtaskResults.find(
            (result) => result.subtaskId === subtask.id,
          );
          return {
            verifierFeedback: subtaskResult?.feedback ?? verification.feedback,
            verifierIssues: subtaskResult?.issues ?? verification.issues,
          };
        },
        stepRetryAttempts,
      );
    }

    // Cancelled means the client is already gone — writing to a closed
    // response would throw for no benefit, since nothing is left to read it.
    if (!verifierPassed && !cancelled) {
      emitter.send({
        type: 'VERIFY_FAIL',
        feedback:
          retryExhaustedMessage ?? 'Could not verify after scoped retries — showing best attempt',
      });
    }

    // Own states only: a stepwise wave reports on the work IT did. Prior waves
    // were already reported, accepted and change-set'd on their own requests —
    // re-emitting them here would double-count every action in the run.
    this.applyVerifyResultsToStates(ownStates, lastSubtaskVerifyResults, verifierPassed);

    return this.buildLoopResult(ownStates, iterationsRun, verifierPassed, {
      preferCompletedOnly: !verifierPassed,
      defaultFailReason: cancelled
        ? 'Cancelled — client disconnected before verification finished'
        : retryExhaustedMessage ?? 'Could not complete and verify the full request',
      timedOut,
    });
  }

  private applyVerifyResultsToStates(
    subtaskStates: SubtaskActionState[],
    results: Array<{ subtaskId: string; passed: boolean; feedback: string }>,
    verifierPassed: boolean,
  ): void {
    if (verifierPassed) {
      for (const state of subtaskStates) {
        state.verified = true;
        // TASKS.md #274 — this used to stamp `completed = true`
        // UNCONDITIONALLY here, which meant a wave-level "verified" verdict
        // (including `verifier-skip.policy.ts` deciding the wave's actions
        // look low-risk enough to skip verification, with NO knowledge of
        // per-subtask completion) could resurrect a subtask that hit its own
        // 10-iteration cap and never finished. By this point `executeSubtask`
        // has ALREADY run to completion for every state in this wave and, for
        // one that never finished, has ALREADY recorded why (hit max
        // iterations / timed out / blocked) in `state.failedReason` — that is
        // ground truth this wave-level verdict must not override. Only a
        // state with no such reason gets stamped, which for a genuinely
        // finished subtask is a no-op (executeSubtask already set it).
        if (!state.failedReason) {
          state.completed = true;
        }
      }
      return;
    }

    for (const result of results) {
      const state = subtaskStates.find((entry) => entry.subtask.id === result.subtaskId);
      if (!state) continue;
      state.verified = result.passed;
      if (result.passed) {
        state.completed = true;
      } else if (!state.failedReason) {
        state.failedReason = result.feedback || 'Verification failed for this step';
      }
    }
  }

  private buildLoopResult(
    subtaskStates: SubtaskActionState[],
    iterationsRun: number,
    verifierPassed: boolean,
    options: {
      preferCompletedOnly: boolean;
      defaultFailReason: string;
      timedOut?: boolean;
    },
  ): AgenticLoopResult {
    const candidateCompleted: CompletedSubtaskResult[] = subtaskStates
      .filter((state) => {
        if (!state.completed || state.actions.length === 0) return false;
        if (verifierPassed) return true;
        // Partial progress: only include steps that verified or finished cleanly before a later failure
        return state.verified === true || (state.completed && !state.failedReason);
      })
      .map((state) => ({
        subtaskId: state.subtask.id,
        actions: state.actions,
        verified: state.verified === true || verifierPassed,
      }));

    // TASKS.md #195 — collect EVERY failed subtask, not just the first. A wave
    // of independent parallel subtasks (no dependsOn between them, e.g. 12
    // month-sheet creates) can have several genuinely fail at once; using
    // `.find()` alone silently discarded every failure but one, with no
    // recorded reason anywhere for the rest — the exact live incident this
    // fixes (9 of 12 month sheets failed, only 1 reason was ever recorded).
    const explicitlyFailedStates = subtaskStates.filter((state) => Boolean(state.failedReason));
    // A subtask with no explicit failedReason but that never verified/completed
    // still failed — same fallback single-pass logic used before, just applied
    // to every such state instead of only the first one found.
    //
    // TASKS.md #274 — no longer gated on `!verifierPassed`. A wave-level
    // "verified" verdict says nothing about whether every INDIVIDUAL subtask
    // in it actually finished (the skip-verifier policy in particular judges
    // the wave's accumulated actions, not per-subtask completion) — so a
    // subtask that hit its own iteration cap must still surface as a failure
    // here, or its actions get silently dropped by the filters below with no
    // reason ever reported to the caller. For a genuinely all-succeeded wave
    // this changes nothing: every state is already `completed` with no
    // `failedReason`, so neither filter matches and this stays empty.
    const implicitlyFailedStates = subtaskStates.filter(
      (state) =>
        !state.failedReason &&
        (state.verified === false || !state.completed) &&
        !explicitlyFailedStates.includes(state),
    );

    let failedSubtasks: FailedSubtaskResult[] = [
      ...explicitlyFailedStates,
      ...implicitlyFailedStates,
    ].map((state) => ({
      subtaskId: state.subtask.id,
      reason:
        state.failedReason ??
        (options.timedOut
          ? 'Timed out before this step completed'
          : options.defaultFailReason),
    }));

    // The single most-relevant failure, for callers/messages that only ever
    // describe one — NOT necessarily the only failure. See its own docblock.
    let failedSubtask: FailedSubtaskResult | null = failedSubtasks[0] ?? null;

    // Spec 22 Bug 2: never ship destructive actions as partial progress when
    // dependencies failed or the full chain did not verify.
    const completedSubtasks = verifierPassed
      ? candidateCompleted
      : this.filterSafePartialDelivery(candidateCompleted, subtaskStates, failedSubtask);

    const withheldDestructive =
      !verifierPassed &&
      candidateCompleted.some((entry) =>
        entry.actions.some((action) => isDestructiveActionType(action.type)),
      ) &&
      !completedSubtasks.some((entry) =>
        entry.actions.some((action) => isDestructiveActionType(action.type)),
      );

    if (withheldDestructive && failedSubtask) {
      const annotated = {
        ...failedSubtask,
        reason: `${failedSubtask.reason} — withheld destructive change(s) until prerequisites succeed`,
      };
      failedSubtask = annotated;
      failedSubtasks = [annotated, ...failedSubtasks.slice(1)];
    } else if (withheldDestructive && !failedSubtask) {
      failedSubtask = {
        subtaskId: candidateCompleted.find((e) =>
          e.actions.some((a) => isDestructiveActionType(a.type)),
        )?.subtaskId ?? 'destructive',
        reason:
          'Withheld destructive change(s) because the full request could not be verified safely',
      };
      failedSubtasks = [failedSubtask];
    }

    // A subtask that finished but was held back by the partial-delivery policy
    // is a failure to report, not a silent drop. TASKS.md #315.
    const heldBack = candidateCompleted.filter(
      (entry) =>
        !completedSubtasks.some((kept) => kept.subtaskId === entry.subtaskId) &&
        !failedSubtasks.some((failed) => failed.subtaskId === entry.subtaskId),
    );
    if (heldBack.length > 0) {
      failedSubtasks = [
        ...failedSubtasks,
        ...heldBack.map((entry) => ({
          subtaskId: entry.subtaskId,
          reason: 'Held back — a step it depends on did not complete',
        })),
      ];
      failedSubtask = failedSubtask ?? failedSubtasks[0];
    }
    const heldBackIds = new Set(heldBack.map((entry) => entry.subtaskId));

    const partialProgress =
      !verifierPassed && completedSubtasks.length > 0 && failedSubtask !== null;

    let actions: Action[];
    if (verifierPassed) {
      // TASKS.md #274 — `verifierPassed` is a WAVE-level verdict (real LLM
      // pass, a clean deterministic pass, OR `verifier-skip.policy.ts`
      // deciding the wave's accumulated actions look "low-risk" enough to
      // skip verification entirely) and says nothing about whether any ONE
      // subtask in the wave actually finished. `state.completed` does: it is
      // set only when that subtask's OWN Executor call returned `isDone`,
      // independent of what the wave-level verdict concludes afterward.
      //
      // The live failure: a single-subtask wave (ADD_SHEET + a headerless
      // CREATE_TABLE) hit the 10-iteration cap and never reached `isDone`.
      // Those 2 actions are individually non-destructive/non-formula, so
      // `shouldSkipVerifier` waved the wave through as "low-risk" — the skip
      // policy has no concept of "did the subtask that produced these
      // actions ever finish". Unfiltered flattening then shipped the
      // headerless table as if the step had succeeded. Same principle as the
      // legacy branch below — prior-wave states are always seeded
      // `completed: true`, so this changes nothing for already-accepted work.
      actions = subtaskStates.filter((state) => state.completed).flatMap((state) => state.actions);
    } else if (
      options.preferCompletedOnly &&
      (partialProgress || withheldDestructive)
    ) {
      // Spec 22: filtered list — empty when only unsafe destructive actions were withheld.
      actions = completedSubtasks.flatMap((entry) => entry.actions);
    } else {
      // TASKS.md #274 — this used to be `this.flattenActions(subtaskStates)`,
      // unconditionally: every subtask's accumulated actions, with NO check
      // that the subtask itself ever finished. That is how a month-sheet
      // subtask that hit the 10-iteration cap midway — after ADD_SHEET and
      // CREATE_TABLE but BEFORE the header BATCH_SET — still shipped a
      // headerless table (Excel's own "Column1…Column13" placeholders) as if
      // the step had succeeded: this branch runs exactly when
      // `completedSubtasks` is empty (nothing safely completed), which is
      // precisely the case where "flatten everything anyway" is most wrong.
      // A parallel wave of 12 independent month-creates can have several die
      // mid-iteration at once — #195 already taught this codebase not to
      // assume a wave's failures are singular — and blindly shipping their
      // fragments is the false-completeness shape CODEBASE_ANALYSIS.md §3.7
      // keeps re-teaching: the user saw "✓ Applied" on a step whose own
      // Executor never finished it.
      //
      // Restricted to subtasks that reached `completed: true` ON THEIR OWN —
      // narrower than `candidateCompleted` above (no verified/failedReason
      // gate), so a single subtask that finished cleanly but the wave's
      // overall multi-subtask chain never verified still surfaces (the
      // "tests / single-step fails" case the old comment named). A subtask
      // that never finished contributes nothing, which for a wave where NO
      // subtask completed means this now correctly ships zero actions instead
      // of a workbook of half-built sheets.
      actions = subtaskStates
        .filter((state) => state.completed && !heldBackIds.has(state.subtask.id))
        .flatMap((state) => state.actions);
    }

    // What is RECORDED as delivered must be exactly what ships. The branches
    // above that flatten `state.completed` ship subtasks `completedSubtasks`
    // may have filtered out, and the stepwise path records completion from
    // `completedSubtasks` — so a live run applied 12 month sheets while
    // recording all 12 as not done, hiding them from every later wave.
    // TASKS.md #318.
    const shipsCompletedStates =
      verifierPassed || !(options.preferCompletedOnly && (partialProgress || withheldDestructive));
    const deliveredSubtasks: CompletedSubtaskResult[] = shipsCompletedStates
      ? subtaskStates
          .filter(
            (state) =>
              state.completed && state.actions.length > 0 && !heldBackIds.has(state.subtask.id),
          )
          .map((state) => ({
            subtaskId: state.subtask.id,
            actions: state.actions,
            verified: state.verified === true || verifierPassed,
          }))
      : completedSubtasks;

    // Parallel subtasks writing the same shared cells: keep one copy, or the
    // client's overwrite guard refuses the second and the step can never be
    // accepted. TASKS.md #317.
    const deduped = dedupeIdenticalWrites(actions);
    if (deduped.removed > 0 || deduped.conflicts.length > 0) {
      this.logger.warn(
        `Wave writes: dropped ${deduped.removed} identical duplicate write(s)` +
          (deduped.conflicts.length > 0
            ? `; ${deduped.conflicts.length} cell(s) written with DIFFERENT values by more than one action: ` +
              deduped.conflicts.slice(0, 5).join(', ')
            : ''),
      );
    }
    actions = deduped.actions;

    return {
      actions,
      iterationsRun,
      verifierPassed,
      completedSubtasks: deliveredSubtasks,
      failedSubtask,
      failedSubtasks,
      partialProgress,
    };
  }

  /**
   * Spec 22: use Planner dependsOn + destructive-type policy so partial delivery
   * never surfaces DELETE_COLUMN / CLEAR_* alone when a prerequisite failed.
   */
  private filterSafePartialDelivery(
    candidates: CompletedSubtaskResult[],
    subtaskStates: SubtaskActionState[],
    failedSubtask: FailedSubtaskResult | null,
  ): CompletedSubtaskResult[] {
    const byId = new Map(subtaskStates.map((s) => [s.subtask.id, s]));
    const passedIds = new Set(
      candidates.filter((c) => c.verified).map((c) => c.subtaskId),
    );
    // Treat cleanly completed (no fail reason) as available deps for non-destructive siblings.
    for (const c of candidates) {
      const state = byId.get(c.subtaskId);
      if (state && state.completed && !state.failedReason) {
        passedIds.add(c.subtaskId);
      }
    }

    const hasUnmetFailure = failedSubtask !== null;

    return candidates.filter((entry) => {
      const state = byId.get(entry.subtaskId);
      const dependsOn = state?.subtask.dependsOn ?? [];
      // A dependency with no state here belongs to an EARLIER stepwise wave —
      // already built and accepted, or wave gating (#261) would not have let
      // this subtask run. Treating it as unmet dropped every finished sibling
      // the moment any peer in the wave failed: 12 built month sheets in one
      // live run, Main's Monthly Totals in another, each recorded as not done
      // with no reason. TASKS.md #315.
      const depsMet = dependsOn.every((depId) => passedIds.has(depId) || !byId.has(depId));
      if (!depsMet) {
        return false;
      }

      const hasDestructive = entry.actions.some((action) =>
        isDestructiveActionType(action.type),
      );
      if (!hasDestructive) {
        return true;
      }

      // Destructive: only deliver when every dependsOn passed AND no peer
      // failure remains (order-dependent compound requests).
      if (hasUnmetFailure) {
        return false;
      }
      return depsMet && dependsOn.every((depId) => {
        const dep = byId.get(depId);
        return dep?.verified === true || (dep?.completed && !dep.failedReason);
      });
    });
  }

  /** Sheet names these states actually CREATED (not merely wrote to). TASKS.md #294. */
  private sheetsCreatedBy(states: SubtaskActionState[]): Set<string> {
    const names = new Set<string>();
    for (const state of states) {
      for (const action of state.actions) {
        if (action.type !== 'ADD_SHEET' && action.type !== 'CREATE_SHEET') continue;
        const record = action as unknown as Record<string, unknown>;
        const raw = String(record.name ?? record.sheetName ?? '').trim();
        if (raw) names.add(raw.toLowerCase());
      }
    }
    return names;
  }

  private runDeterministicChecks(
    originalPrompt: string,
    subtasks: SubTask[],
    subtaskStates: SubtaskActionState[],
    context: WorkbookContext,
    /**
     * The workbook as it was BEFORE this run — deliberately separate from
     * `context`, which is enriched from the shadow workbook. TASKS.md #294.
     */
    preRunContext: WorkbookContext = context,
    /** Sheets earlier waves of this run already created — TASKS.md #294. */
    sheetsCreatedByEarlierWaves: Set<string> = new Set(),
    /**
     * The workbook as it stood before the subtasks being graded wrote
     * anything — earlier waves included, their own actions excluded.
     * TASKS.md #302.
     */
    contextBeforeOwnWrites: WorkbookContext = context,
  ): CheckerResult {
    const completeness = this.completenessChecker.check(subtasks, subtaskStates);
    const formatting = this.formattingChecker.check(subtaskStates, context);
    const semantic = this.semanticFormulaChecker.check(
      originalPrompt,
      subtasks,
      subtaskStates,
      context,
    );
    // TASKS.md #302 — MUST be the context WITHOUT these subtasks’ own
    // writes. Every other checker here wants to see the result of the work;
    // this one asks whether the cell was occupied BEFORE it, and handing it
    // the shadow-enriched context makes every write self-refuting. The same
    // distinction #294 had to draw, in a different guard.
    //
    // Narrowing worth stating: two subtasks in the SAME wave writing the
    // same cell are no longer caught here. The frontend’s
    // `guardAgainstOverwrite` still runs against the real workbook at apply
    // time and is the authoritative guard; this checker exists to catch it
    // earlier, not to be the only one that does.
    const overwriteOccupancy = this.overwriteOccupancyChecker.check(
      subtaskStates,
      contextBeforeOwnWrites,
    );
    // TASKS.md #294 — MUST be the pre-run context, not the shadow-enriched
    // one. `virtualApply`'s `ensureSheet` conjures a sheet the moment anything
    // writes to it, so a subtask that wrote to "Main" without ever emitting
    // ADD_SHEET made Main appear in the shadow — and this checker's "skip when
    // the sheet already existed" guard (correct for genuinely pre-existing
    // sheets) then skipped the very check that exists to catch that. Live
    // result: a run reported every step complete with Main written to but
    // never created.
    const structuralIntent = this.structuralIntentChecker.check(
      subtaskStates,
      preRunContext,
      sheetsCreatedByEarlierWaves,
    );
    // TASKS.md #270 — a template whose computed columns never got a formula.
    const computedColumn = this.computedColumnChecker.check(subtaskStates);
    // LONG_PROMPT_RELIABILITY_PLAN.md Phase 1 — header row must match the user's column list.
    const specConformance = this.specConformanceChecker.check(subtaskStates);
    const merged = mergeCheckerResults([
      completeness,
      formatting,
      semantic,
      overwriteOccupancy,
      structuralIntent,
      computedColumn,
      specConformance,
    ]);

    const needsSemanticReview =
      subtaskStates.some((state) =>
        state.actions.some(
          (action) =>
            action.type === 'SET_FORMULA' ||
            (typeof action.formula === 'string' && action.formula.startsWith('=')),
        ),
      ) ||
      subtasks.some((subtask) => subtask.dependsOn.length > 0) ||
      subtasks.length > 2;

    return {
      ...merged,
      requiresLlmVerification: merged.requiresLlmVerification || needsSemanticReview,
    };
  }

  private async retrySubtasks(
    subtasksToRetry: SubTask[],
    subtaskStates: SubtaskActionState[],
    context: WorkbookContext,
    emitter: SseEmitter,
    startedAt: number,
    formulaValidationLog: FormulaValidationResult[],
    loopOptions: AgenticLoopOptions,
    failingIds: Set<string>,
    checkResult: CheckerResult,
    onTimeout: () => void,
    originalPrompt: string,
    retryContextFor?: (subtask: SubTask) => RetryContext,
    stepRetryAttempts?: Map<string, number>,
  ): Promise<number> {
    let iterationsRun = 0;
    const completedIds = new Set(
      subtaskStates
        .map((state) => state.subtask.id)
        .filter((id) => !failingIds.has(id)),
    );

    for (const subtask of subtasksToRetry) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        onTimeout();
        break;
      }

      const state = subtaskStates.find((entry) => entry.subtask.id === subtask.id);
      if (!state) continue;

      const subtaskResult = checkResult.subtaskResults.find(
        (result) => result.subtaskId === subtask.id,
      );
      const feedback = subtaskResult?.feedback ?? checkResult.feedback;
      const issues: VerifierIssue[] = subtaskResult?.issues ?? checkResult.issues;

      state.actions = [];
      state.droppedActions = [];
      state.completed = false;
      state.verified = undefined;
      state.failedReason = undefined;
      const visibleIds = new Set([...completedIds, ...subtask.dependsOn]);
      const retryAttempt = (stepRetryAttempts?.get(subtask.id) ?? 0) + 1;
      stepRetryAttempts?.set(subtask.id, retryAttempt);
      iterationsRun += await this.executeSubtask(
        state,
        subtaskStates,
        context,
        emitter,
        startedAt,
        formulaValidationLog,
        loopOptions,
        visibleIds,
        onTimeout,
        originalPrompt,
        retryContextFor?.(subtask) ?? { verifierFeedback: feedback, verifierIssues: issues },
        retryAttempt,
      );
    }

    return iterationsRun;
  }

  /**
   * Runs `worker` over `items` with at most `limit` in flight at once,
   * resolving once every item has settled (never rejects — `worker` is
   * expected to catch its own errors, same contract the wave-execution
   * caller already relies on). Results are returned in `items` order,
   * regardless of completion order. TASKS.md #275.
   */
  /**
   * Same contract as `runWithConcurrencyLimit`, but the width is re-read from
   * an `AdaptiveConcurrency` controller before every batch and updated from
   * that batch's own outcomes — LONG_PROMPT_RELIABILITY_PLAN.md Phase 3
   * (TASKS.md #286). A provider that starts pushing back narrows the next
   * batch immediately; a clean run earns width back a slot at a time.
   *
   * Batched rather than a rolling pool on purpose: a wave here is a set of
   * near-identical subtasks (twelve month sheets), so batch granularity costs
   * very little, and the width rule stays something you can actually reason
   * about — which a dynamically-gated rolling pool would not.
   */
  private async runWithAdaptiveConcurrency<T, R>(
    items: T[],
    controller: AdaptiveConcurrency,
    worker: (item: T) => Promise<R>,
    classify: (result: R) => { transient: boolean; reason: string } | null,
  ): Promise<R[]> {
    const results: R[] = [];
    let index = 0;

    while (index < items.length) {
      const width = Math.max(1, Math.min(controller.limit, items.length - index));
      const batch = items.slice(index, index + width);
      index += width;

      const settled = await Promise.all(batch.map((item) => worker(item)));
      results.push(...settled);

      let sawTransient = false;
      for (const result of settled) {
        const verdict = classify(result);
        if (verdict?.transient) {
          sawTransient = true;
          this.logger.warn(
            `Adaptive concurrency: transient provider fault (${verdict.reason}) — ` +
              `narrowing from ${controller.limit}.`,
          );
        }
      }
      if (sawTransient) {
        controller.recordTransientFailure();
      } else {
        for (const _ of settled) controller.recordSuccess();
      }
    }

    return results;
  }

  private async runWithConcurrencyLimit<T, R>(
    items: T[],
    limit: number,
    worker: (item: T) => Promise<R>,
  ): Promise<R[]> {
    const results: R[] = new Array(items.length);
    let nextIndex = 0;

    const runNext = async (): Promise<void> => {
      const currentIndex = nextIndex;
      nextIndex += 1;
      if (currentIndex >= items.length) {
        return;
      }
      results[currentIndex] = await worker(items[currentIndex]);
      await runNext();
    };

    const workerCount = Math.max(1, Math.min(limit, items.length));
    await Promise.all(Array.from({ length: workerCount }, () => runNext()));
    return results;
  }

  private async executeSubtask(
    state: SubtaskActionState,
    allStates: SubtaskActionState[],
    baseContext: WorkbookContext,
    emitter: SseEmitter,
    startedAt: number,
    formulaValidationLog: FormulaValidationResult[],
    loopOptions: AgenticLoopOptions,
    visibleStateIds: Set<string>,
    onTimeout: () => void,
    originalPrompt: string,
    retryContext?: RetryContext,
    retryAttempt?: number,
  ): Promise<number> {
    const { subtask } = state;
    let iterationsRun = 0;

    this.logger.log(`Agentic loop: starting subtask "${subtask.description}"`);
    emitter.send({ type: 'CHECKPOINT', step: subtask.description });

    let iteration = 0;
    let subtaskDone = false;

    while (!subtaskDone && iteration < this.MAX_ITERATIONS_PER_SUBTASK) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        onTimeout();
        break;
      }

      // TASKS.md #281 — abortSignal was only ever checked BETWEEN waves and
      // in the verify/retry cycle, never inside this per-subtask iteration
      // loop. A live run showed exactly what that costs: Stop did nothing
      // visible while a single heavy subtask ground through its own retries —
      // the client's fetch was aborted, but the server kept calling the
      // Executor for every remaining iteration regardless, on a connection
      // nobody was reading anymore. Checked every iteration, same as the
      // wall-clock check right above it.
      if (loopOptions.abortSignal?.aborted) {
        state.failedReason = 'Cancelled — client disconnected';
        this.logger.warn(
          `Subtask "${subtask.description}" cancelled mid-iteration (client disconnected/stopped)`,
        );
        break;
      }

      iteration += 1;
      iterationsRun += 1;

      emitter.send({
        type: 'THINKING',
        message:
          iteration > 1
            ? `Continuing "${subtask.description}" (step ${iteration})...`
            : `Working on "${subtask.description}"...`,
      });

      const shadow = this.buildShadowFromStates(baseContext, allStates, visibleStateIds);
      const relevantSheetNames = this.resolveSubtaskRelevantSheets(subtask, baseContext);
      const currentContext = {
        ...this.enrichContextFromShadow(shadow, relevantSheetNames),
        ...retryContext,
      };
      const previousActions = this.flattenActions(
        allStates.filter((entry) => visibleStateIds.has(entry.subtask.id)),
      );

      const validatedBatch = await this.runExecutorWithFormulaValidation(
        subtask,
        currentContext,
        shadow,
        previousActions,
        emitter,
        startedAt,
        onTimeout,
        formulaValidationLog,
        loopOptions,
        retryAttempt && retryContext?.verifierFeedback
          ? {
              originalStep: subtask,
              attempt: retryAttempt,
              maxAttempts: this.MAX_STEP_RETRIES,
              verifierFeedback: retryContext.verifierFeedback,
            }
          : undefined,
      );

      if (!validatedBatch) {
        emitter.send({
          type: 'THINKING',
          message: `Formula validation blocked actions for "${subtask.description}" — retrying`,
        });
        continue;
      }

      // Never invent date/number display formats — rebind to sheet-owned formats or drop.
      const rebound = rebindFormatRangeNumberFormats(validatedBatch.actions, {
        userPrompt: originalPrompt,
        subtaskDescription: subtask.description,
        context: currentContext,
      });
      if (rebound.droppedInvented) {
        this.logger.log(
          `Stripped invented numberFormat on "${subtask.description}" (user did not name a format)`,
        );
      }
      if (rebound.actions.length === 0 && validatedBatch.actions.length > 0) {
        state.failedReason =
          'Could not apply a number format without inventing one — please name the format (e.g. m/d/yyyy) or ask to re-use the existing column format.';
        emitter.send({
          type: 'THINKING',
          message: state.failedReason,
        });
        break;
      }

      for (const action of rebound.actions) {
        emitter.send({ type: 'ACTION', action });
        state.actions.push(action);
      }
      if (validatedBatch.droppedActions?.length) {
        state.droppedActions.push(...validatedBatch.droppedActions);
      }

      subtaskDone = validatedBatch.isDone;

      if (
        !subtaskDone &&
        validatedBatch.actions.length === 0 &&
        isExecutorBlockedSignal(validatedBatch.nextStep)
      ) {
        state.failedReason = validatedBatch.nextStep;
        this.logger.warn(
          `Executor blocked on "${subtask.description}": ${validatedBatch.nextStep}`,
        );
        emitter.send({
          type: 'THINKING',
          message: validatedBatch.nextStep ?? 'Blocked — cannot complete this step',
        });
        break;
      }

      if (!subtaskDone && validatedBatch.nextStep) {
        emitter.send({ type: 'THINKING', message: validatedBatch.nextStep });
      }
    }

    if (subtaskDone) {
      state.completed = true;
      state.failedReason = undefined;
    } else {
      // Preserve an honest block reason — do not overwrite with "max iterations".
      if (!state.failedReason) {
        const hitTimeout = Date.now() - startedAt > this.TIMEOUT_MS;
        state.failedReason = hitTimeout
          ? `Subtask "${subtask.description}" timed out before completion`
          : `Subtask "${subtask.description}" hit max iterations (${this.MAX_ITERATIONS_PER_SUBTASK})`;
        this.logger.warn(state.failedReason);
        emitter.send({
          type: 'THINKING',
          message: hitTimeout
            ? `Timed out on "${subtask.description}" — continuing with work already ready`
            : `Reached step limit for "${subtask.description}" — moving on`,
        });
      } else {
        this.logger.warn(
          `Subtask "${subtask.description}" stopped: ${state.failedReason}`,
        );
        emitter.send({
          type: 'THINKING',
          message: state.failedReason,
        });
      }
    }

    return iterationsRun;
  }

  /** Run executor with optional tool fetch, pre/post formula validation, and retries. */
  private async runExecutorWithFormulaValidation(
    subtask: SubTask,
    context: WorkbookContext,
    baseShadow: ShadowWorkbook,
    previousActions: Action[],
    emitter: SseEmitter,
    startedAt: number,
    onTimeout: () => void,
    formulaValidationLog: FormulaValidationResult[],
    loopOptions: AgenticLoopOptions,
    retryStepContext?: StepRetryContext,
  ): Promise<ExecutorOutput | null> {
    const callStartedAt = Date.now();
    const model = this.executor.modelName;
    let execContext = { ...context };
    let result = buildDeterministicSubtaskActions(subtask, execContext);
    if (result) {
      this.logger.log(
        `Using ${result.actions.length} deterministic action(s) for "${subtask.description}"`,
      );
    } else {
      try {
        result = retryStepContext
          ? await this.executor.retryStep(
              retryStepContext,
              execContext,
              previousActions,
              loopOptions.correlationId,
              loopOptions.usageTotals,
            )
          : await this.executor.execute(
              subtask,
              execContext,
              previousActions,
              loopOptions.correlationId,
              loopOptions.usageTotals,
            );
        this.noteExecutorParseResult(result, loopOptions);
      } catch (error) {
        if (error instanceof StepRetryExhaustedError) {
          this.logger.error(error.message);
          this.structuredLogger.logAgentEvent({
            correlationId: loopOptions.correlationId ?? '-',
            agent: 'workbook',
            model,
            durationMs: Date.now() - callStartedAt,
            success: false,
            error: error.message,
          });
          return null;
        }
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: false,
          error: error instanceof Error ? error.message : String(error),
        });
        throw error;
      }
    }

    let toolAttempts = 0;
    while (
      result.toolRequest &&
      toolAttempts < this.MAX_TOOL_REQUESTS &&
      loopOptions.conversationId &&
      loopOptions.toolEmit
    ) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        onTimeout();
        return null;
      }

      toolAttempts += 1;
      const { toolRequest } = result;
      emitter.send({
        type: 'THINKING',
        message: `Fetching range ${toolRequest.range} on ${toolRequest.sheet}...`,
      });

      try {
        const fetched = await this.toolBridge.waitForRangeData(
          loopOptions.conversationId,
          toolRequest,
          loopOptions.toolEmit,
        );

        if (fetched.error) {
          emitter.send({
            type: 'THINKING',
            message: `Range fetch failed: ${fetched.error}`,
          });
          return null;
        }

        execContext = this.mergeFetchedRange(execContext, toolRequest, fetched.values);
        result = await this.executor.execute(
          subtask,
          execContext,
          previousActions,
          loopOptions.correlationId,
          loopOptions.usageTotals,
        );
        this.noteExecutorParseResult(result, loopOptions);
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Range fetch failed';
        this.logger.warn(`Tool request failed: ${message}`);
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: false,
          error: message,
        });
        emitter.send({ type: 'THINKING', message });
        return null;
      }
    }

    if (result.toolRequest && result.actions.length === 0) {
      emitter.send({
        type: 'THINKING',
        message: 'Could not fetch requested range — try a smaller range or simpler task',
      });
      return null;
    }

    let formulaAttempt = 0;

    while (formulaAttempt <= this.MAX_FORMULA_RETRIES) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        onTimeout();
        return null;
      }

      const preValidation = this.formulaValidator.validatePreApply(
        result.actions,
        execContext,
        subtask.targetSheet,
        baseShadow,
      );
      formulaValidationLog.push(preValidation);

      if (!preValidation.passed) {
      if (formulaAttempt >= this.MAX_FORMULA_RETRIES) {
        const deterministic = buildDeterministicSubtaskActions(subtask, execContext);
        if (deterministic?.actions.length) {
          this.logger.log(
            `Formula pre-validation failed — using ${deterministic.actions.length} deterministic action(s)`,
          );
          result = deterministic;
          formulaAttempt = 0;
          continue;
        }
        this.logger.warn(
          `Formula pre-validation failed after ${this.MAX_FORMULA_RETRIES} retries for "${subtask.description}"`,
        );
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: false,
          error: this.formulaValidator.formatFeedback(preValidation.issues),
        });
        emitter.send({
          type: 'THINKING',
          message: `Formula errors remain: ${this.formulaValidator.formatFeedback(preValidation.issues)}`,
        });
        return null;
      }

        formulaAttempt += 1;
        execContext = {
          ...execContext,
          formulaValidationFeedback: this.formulaValidator.formatFeedback(preValidation.issues),
          formulaValidationIssues: preValidation.issues,
        };
        result = await this.executor.execute(
          subtask,
          execContext,
          previousActions,
          loopOptions.correlationId,
          loopOptions.usageTotals,
        );
        this.noteExecutorParseResult(result, loopOptions);
        continue;
      }

      if (result.actions.length === 0) {
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: true,
          parsedResponse: result,
        });
        return result;
      }

      const postShadow = virtualApply(baseShadow, result.actions);
      const postValidation = this.formulaValidator.checkPostApply(
        postShadow,
        result.actions,
        execContext,
        subtask.targetSheet,
      );
      formulaValidationLog.push(postValidation);

      if (postValidation.passed) {
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: true,
          parsedResponse: result,
        });
        return result;
      }

      if (formulaAttempt >= this.MAX_FORMULA_RETRIES) {
        const deterministic = buildDeterministicSubtaskActions(subtask, execContext);
        if (deterministic?.actions.length) {
          this.logger.log(
            `Formula post-validation failed — using ${deterministic.actions.length} deterministic action(s)`,
          );
          result = deterministic;
          formulaAttempt = 0;
          continue;
        }
        this.logger.warn(
          `Formula post-validation failed after ${this.MAX_FORMULA_RETRIES} retries for "${subtask.description}"`,
        );
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: false,
          error: this.formulaValidator.formatFeedback(postValidation.issues),
        });
        emitter.send({
          type: 'THINKING',
          message: `Post-apply formula errors: ${this.formulaValidator.formatFeedback(postValidation.issues)}`,
        });
        return null;
      }

      formulaAttempt += 1;
      execContext = {
        ...execContext,
        formulaValidationFeedback: this.formulaValidator.formatFeedback(postValidation.issues),
        formulaValidationIssues: postValidation.issues,
      };
      result = await this.executor.execute(
        subtask,
        execContext,
        previousActions,
        loopOptions.correlationId,
        loopOptions.usageTotals,
      );
    }

    this.structuredLogger.logAgentEvent({
      correlationId: loopOptions.correlationId ?? '-',
      agent: 'workbook',
      model,
      durationMs: Date.now() - callStartedAt,
      success: false,
      error: `Executor returned null for subtask ${subtask.id}`,
    });
    return null;
  }

  private mergeFetchedRange(
    context: WorkbookContext,
    toolRequest: NonNullable<ExecutorOutput['toolRequest']>,
    values: unknown[][],
  ): WorkbookContext {
    const sheets = context.sheets.map((sheet) => {
      if (sheet.name !== toolRequest.sheet) return sheet;
      return mergeRangeIntoSheet(sheet, toolRequest.range, values);
    });

    const fetchedRanges = [
      ...(context.fetchedRanges ?? []),
      {
        sheet: toolRequest.sheet,
        range: toolRequest.range,
        rowCount: values.length,
      },
    ];

    return {
      ...context,
      sheets,
      fetchedRanges,
    };
  }

  /**
   * Perf #71: buildShadowFromStates() previously rebuilt the ENTIRE shadow from
   * scratch on every call — one call per verifier cycle (agenticLoop's outer
   * while loop) AND one call per Executor iteration within a subtask (the inner
   * while loop at runSubtaskExecution). A 19-subtask batch with 3 verification
   * cycles replayed all 19 subtasks' actions through virtualApply() up to 3
   * times each — most of that work is identical every time, since only the
   * most-recently-changed subtask's actions actually differ between calls.
   *
   * virtualApply() is pure (deep-clones its input, never mutates) — confirmed
   * by reading virtualApply.ts before relying on this — so caching intermediate
   * shadow snapshots by state-array identity is safe: no call site can observe
   * a cached shadow being mutated out from under it.
   *
   * Cache keyed on the array reference of `states` + `visibleStateIds` + the
   * count of already-applied states, walking forward from the longest matching
   * cached prefix rather than replaying from an empty shadow every time. Scoped
   * per AgenticLoopService instance is safe ONLY because this service has no
   * other per-request mutable state and each `run()` call constructs a fresh
   * `subtaskStates` array — different requests never share a states array
   * identity, so entries naturally stop matching and fall out of relevance
   * (bounded further by the small WeakMap-based cache below, GC'd once the
   * states array itself is no longer referenced).
   */
  private readonly shadowPrefixCache = new WeakMap<
    SubtaskActionState[],
    { actionsSnapshot: { ref: Action[]; len: number }[]; shadow: ShadowWorkbook }[]
  >();

  private buildShadowFromStates(
    baseContext: WorkbookContext,
    states: SubtaskActionState[],
    visibleStateIds?: Set<string>,
  ): ShadowWorkbook {
    const relevantStates = visibleStateIds
      ? states.filter((s) => visibleStateIds.has(s.subtask.id))
      : states;

    // Only states with actions actually mutate the shadow — matches the
    // original loop's `if (state.actions.length === 0) continue`.
    const activeActions = relevantStates
      .map((s) => s.actions)
      .filter((actions) => actions.length > 0);

    let cacheEntries = this.shadowPrefixCache.get(states);
    if (!cacheEntries) {
      cacheEntries = [];
      this.shadowPrefixCache.set(states, cacheEntries);
    }

    // Find the longest cached prefix whose (reference, length) pairs still
    // match the corresponding prefix of activeActions. Reference identity
    // alone is NOT sufficient: runSubtaskExecution's inner iteration loop
    // does `state.actions.push(action)` on the SAME array across multiple
    // Executor calls within one subtask (confirmed by reading that call
    // site) — the reference stays stable while content grows, so length is
    // checked alongside reference to catch that in-place-growth case. A
    // subtask RETRY (as opposed to continued iteration) always starts from
    // `state.actions = []`, a fresh reference, so that case is still caught
    // by the reference check regardless of length.
    let startIndex = 0;
    let shadow = buildShadowWorkbook(baseContext);
    for (const entry of cacheEntries) {
      const len = entry.actionsSnapshot.length;
      if (len > activeActions.length || len <= startIndex) continue;
      const matches = entry.actionsSnapshot.every(
        (snap, i) => snap.ref === activeActions[i] && snap.len === activeActions[i].length,
      );
      if (matches) {
        startIndex = len;
        shadow = entry.shadow;
      }
    }

    for (let i = startIndex; i < activeActions.length; i += 1) {
      shadow = virtualApply(shadow, activeActions[i]);
    }

    // Cache the full-prefix result for future calls. Cap growth: keep only
    // the most recent few prefixes (verifier cycles are bounded by
    // maxVerifierCycles, subtask iterations by MAX_ITERATIONS_PER_SUBTASK —
    // neither is large, so an unbounded cache here would still be small, but
    // capping keeps memory flat instead of growing with cycle count).
    cacheEntries.push({
      actionsSnapshot: activeActions.map((ref) => ({ ref, len: ref.length })),
      shadow,
    });
    if (cacheEntries.length > 8) {
      cacheEntries.shift();
    }

    return shadow;
  }

  /**
   * `relevantSheetNames`, when provided, scopes the (relatively expensive,
   * full-formula-walk) analyzeSheet() call to just those sheets — used by the
   * per-subtask Executor context (line ~827) where only the subtask's own
   * target sheet matters. Left undefined (all sheets analyzed) for the
   * whole-batch Verifier context (line ~252), which legitimately needs
   * cross-sheet visibility to catch things like "a dashboard chart pointing
   * at a sheet nobody actually touched" — narrowing that path risks a real
   * regression in verification coverage, not just a perf change.
   */
  private enrichContextFromShadow(
    shadow: ShadowWorkbook,
    relevantSheetNames?: Set<string>,
  ): WorkbookContext {
    const context = shadowToWorkbookContext(shadow);
    return {
      ...context,
      sheets: context.sheets.map((sheet) =>
        !relevantSheetNames || relevantSheetNames.has(sheet.name)
          ? { ...sheet, formulaInsights: this.formulaAnalyzer.analyzeSheet(sheet) }
          : sheet,
      ),
    };
  }

  /**
   * Perf #71 (companion to #70's identical scoping in conversation.service.ts):
   * an Executor call for one subtask only needs formula insight for that
   * subtask's own target sheet, plus any other sheet its description names
   * (covers cross-sheet formula subtasks, e.g. the Monthly Totals table
   * writing SUMIF formulas that reference each month sheet by name). Every
   * OTHER sheet in a large workbook (e.g. 11 other month sheets) previously
   * paid analyzeSheet()'s full-formula-walk cost on every Executor iteration
   * for no reason — nothing in that subtask's own prompt ever reads it.
   */
  private resolveSubtaskRelevantSheets(
    subtask: SubTask,
    context: WorkbookContext,
  ): Set<string> {
    const relevant = new Set<string>([subtask.targetSheet]);
    for (const sheet of context.sheets) {
      if (sheet.name && subtask.description.includes(sheet.name)) {
        relevant.add(sheet.name);
      }
    }
    return relevant;
  }

  private flattenActions(states: SubtaskActionState[]): Action[] {
    return states.flatMap((state) => state.actions);
  }

  private actionsBySubtaskMap(states: SubtaskActionState[]): Record<string, Action[]> {
    return Object.fromEntries(states.map((state) => [state.subtask.id, state.actions]));
  }

  private collectDownstreamSubtasks(
    ordered: SubTask[],
    failingIds: Set<string>,
  ): Set<string> {
    const downstream = new Set<string>(failingIds);

    let changed = true;
    while (changed) {
      changed = false;
      for (const subtask of ordered) {
        if (downstream.has(subtask.id)) continue;
        if (subtask.dependsOn.some((dep) => downstream.has(dep))) {
          downstream.add(subtask.id);
          changed = true;
        }
      }
    }

    return downstream;
  }

  private lockPassedSubtasks(
    results: Array<{ subtaskId: string; passed: boolean; inconclusive?: boolean }>,
    lockedPassIds: Set<string>,
  ): void {
    for (const result of results) {
      if (result.passed && !result.inconclusive) {
        lockedPassIds.add(result.subtaskId);
      }
    }
  }

  /**
   * Spec 17 Bug B: only genuinely failed (not inconclusive, not already locked)
   * subtasks are re-executed.
   */
  private collectFailingIdsForRetry(
    results: Array<{ subtaskId: string; passed: boolean; inconclusive?: boolean }>,
    lockedPassIds: Set<string>,
  ): Set<string> {
    const failing = new Set<string>();
    for (const result of results) {
      if (lockedPassIds.has(result.subtaskId)) continue;
      if (result.inconclusive) continue;
      if (!result.passed) failing.add(result.subtaskId);
    }
    return failing;
  }

  private mergeWithLockedPasses(
    partial: VerifierOutput,
    ordered: SubTask[],
    lockedPassIds: Set<string>,
    previous: Array<{ subtaskId: string; passed: boolean; feedback: string; inconclusive?: boolean }>,
  ): VerifierOutput {
    const byId = new Map(partial.subtaskResults.map((r) => [r.subtaskId, r]));
    const subtaskResults = ordered.map((subtask) => {
      if (lockedPassIds.has(subtask.id)) {
        const prev = previous.find((r) => r.subtaskId === subtask.id);
        return {
          subtaskId: subtask.id,
          passed: true,
          feedback: prev?.feedback ?? 'Previously verified',
          issues: [],
        };
      }
      return (
        byId.get(subtask.id) ?? {
          subtaskId: subtask.id,
          passed: false,
          feedback: 'Missing verifier result',
          issues: [],
          inconclusive: true,
        }
      );
    });

    const passed = subtaskResults.every((r) => r.passed && !r.inconclusive);
    return {
      passed,
      feedback: partial.feedback,
      issues: partial.issues,
      subtaskResults,
    };
  }

  private findExhaustedSubtasks(
    failingIds: Set<string>,
    stepRetryAttempts: Map<string, number>,
  ): Set<string> {
    const exhausted = new Set<string>();
    for (const subtaskId of failingIds) {
      if ((stepRetryAttempts.get(subtaskId) ?? 0) >= this.MAX_STEP_RETRIES) {
        exhausted.add(subtaskId);
      }
    }
    return exhausted;
  }

  private buildRetryExhaustedMessage(exhaustedIds: Set<string>, feedback: string): string {
    const ids = Array.from(exhaustedIds).join(', ');
    return `Could not complete step(s) ${ids} after ${this.MAX_STEP_RETRIES} attempts. ${feedback}`;
  }

  private noteExecutorParseResult(
    result: ExecutorOutput | null | undefined,
    loopOptions: AgenticLoopOptions,
  ): void {
    if (result?.parsedOnFirstAttempt === false && loopOptions.parseFailureTracker) {
      loopOptions.parseFailureTracker.hadFailure = true;
    }
  }

  private orderByDependencies(subtasks: SubTask[]): SubTask[] {
    const byId = new Map(subtasks.map((subtask) => [subtask.id, subtask]));
    const visited = new Set<string>();
    const result: SubTask[] = [];

    const visit = (id: string) => {
      if (visited.has(id)) return;
      const task = byId.get(id);
      if (!task) return;
      for (const dep of task.dependsOn) {
        visit(dep);
      }
      visited.add(id);
      result.push(task);
    };

    for (const task of subtasks) {
      visit(task.id);
    }

    return result;
  }
}
