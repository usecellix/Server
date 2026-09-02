import { Injectable, Logger } from '@nestjs/common';
import { PlannerAgent } from './planner.agent';
import { AgenticLoopService } from './agenticLoop.service';
import { SseEmitter } from './sse.emitter';
import { Action, AgentRunOptions, PlannerOutput } from './types/agent.types';
import { annotateExplicitOverwriteConfirmation } from '../excel-ai/utils/overwrite-confirmation.util';
import { annotateClearIntentOverwrite } from './utils/clear-intent-overwrite.util';
import { pruneSpuriousAddSheetActions } from './utils/compound-action.util';
import { createUsageAccumulator } from './utils/usage-accumulator.util';
import { LlmCallTelemetry } from '../excel-ai/services/openrouter.service';

export interface OrchestratorRunResult {
  actions: Action[];
  iterationsRun: number;
  verifierPassed: boolean;
  clarificationRequested: boolean;
  completedSubtasks: Array<{ subtaskId: string; actions: Action[]; verified: boolean }>;
  failedSubtask: { subtaskId: string; reason: string } | null;
  partialProgress: boolean;
  /**
   * The plan's own natural-language statements of intent, surfaced so the
   * Accept card can describe what the build DOES rather than enumerate the
   * mechanical actions it emits. The Planner already writes exactly the right
   * text ("Create sheet 'January' ... and set A1:J1 headers [...]"); it was
   * previously streamed as a transient `status` event and then discarded,
   * leaving the card to render ~40 action bullets instead. TASKS.md #149.
   */
  planSubtasks: Array<{ id: string; description: string; targetSheet: string }>;
  /**
   * Subtasks the Planner produced that the Executor delivered no actions for.
   * Non-empty means the build is incomplete relative to its own plan — the
   * `estimatedActions` circularity `CODEBASE_ANALYSIS.md` §3.15 describes,
   * caught here at the one place both numbers are known. TASKS.md #155.
   */
  undeliveredSubtasks: Array<{ id: string; description: string; targetSheet: string }>;
  /**
   * Questions the Planner raised that did NOT block the build.
   *
   * Non-empty means the plan was produced under a stated assumption (a default
   * year, a placeholder dropdown list). Surfaced with the result so the user
   * can correct it, rather than swallowed. TASKS.md #171.
   */
  openQuestions: string[];
}

@Injectable()
export class OrchestratorService {
  private readonly logger = new Logger(OrchestratorService.name);

  constructor(
    private readonly planner: PlannerAgent,
    private readonly agenticLoop: AgenticLoopService,
  ) {}

  /**
   * Plan mode: run only the PlannerAgent and return its structured plan without
   * executing any actions against the workbook.
   *
   * `telemetry` is an out-param (same pattern as `LlmCompletionOutcome`) — when
   * provided, filled in with the real usage/model this call made, so the
   * caller's audit log can report actual tokens instead of the 0 it always
   * reported before this existed.
   */
  async planOnly(opts: AgentRunOptions, telemetry?: LlmCallTelemetry): Promise<PlannerOutput> {
    const {
      prompt,
      context,
      conversationHistory = [],
      promptContext,
      correlationId,
      routerAssumption,
      complexity,
    } = opts;
    const resolvedCorrelationId = this.resolveCorrelationId(correlationId);
    const usageTotals = createUsageAccumulator();
    try {
      return await this.planner.plan(
        prompt,
        context,
        conversationHistory,
        promptContext,
        resolvedCorrelationId,
        routerAssumption,
        complexity,
        usageTotals,
      );
    } finally {
      this.applyUsageToTelemetry(telemetry, usageTotals);
    }
  }

  async run(opts: AgentRunOptions, emitter: SseEmitter): Promise<Action[]> {
    const result = await this.runDetailed(opts, emitter);
    return result.actions;
  }

  /**
   * `telemetry` is an out-param (same pattern as `LlmCompletionOutcome`) —
   * when provided, filled in with the real usage/model accumulated across the
   * Planner AND every Executor/Verifier call this run makes, so the caller's
   * audit log can report actual tokens instead of the 0 it always reported
   * before this existed. Populated in a `finally` so a run that throws (or
   * takes the early CLARIFY return) still reports whatever usage it burned
   * before failing, rather than silently losing it.
   */
  async runDetailed(
    opts: AgentRunOptions,
    emitter: SseEmitter,
    telemetry?: LlmCallTelemetry,
  ): Promise<OrchestratorRunResult> {
    const {
      prompt,
      context,
      conversationHistory = [],
      promptContext,
      conversationId,
      correlationId,
      toolEmit,
      routerAssumption,
      complexity,
    } = opts;
    const resolvedCorrelationId = this.resolveCorrelationId(correlationId);
    const usageTotals = createUsageAccumulator();

    try {
      return await this.runDetailedWithUsage(
        prompt,
        context,
        conversationHistory,
        promptContext,
        conversationId,
        toolEmit,
        routerAssumption,
        complexity,
        resolvedCorrelationId,
        emitter,
        usageTotals,
      );
    } finally {
      this.applyUsageToTelemetry(telemetry, usageTotals);
    }
  }

  private async runDetailedWithUsage(
    prompt: string,
    context: AgentRunOptions['context'],
    conversationHistory: NonNullable<AgentRunOptions['conversationHistory']>,
    promptContext: AgentRunOptions['promptContext'],
    conversationId: AgentRunOptions['conversationId'],
    toolEmit: AgentRunOptions['toolEmit'],
    routerAssumption: AgentRunOptions['routerAssumption'],
    complexity: AgentRunOptions['complexity'],
    resolvedCorrelationId: string,
    emitter: SseEmitter,
    usageTotals: ReturnType<typeof createUsageAccumulator>,
  ): Promise<OrchestratorRunResult> {
    emitter.send({ type: 'THINKING', message: 'Planning your request...' });
    const plan: PlannerOutput = await this.planner.plan(
      prompt,
      context,
      conversationHistory,
      promptContext,
      resolvedCorrelationId,
      routerAssumption,
      complexity,
      usageTotals,
    );

    // Block ONLY when there is nothing to build — TASKS.md #171.
    //
    // This branch used to bail on `clarificationsNeeded.length > 0` alone, and a
    // live run showed what that costs: the Planner returned a COMPLETE
    // 23-subtask plan (sheets, tables, formulas, dropdowns, dashboard) together
    // with two questions — which year, and what the real dropdown values are —
    // and the whole plan was thrown away to ask them. The user got a dead end
    // where a workbook was one step from existing.
    //
    // Worse, the questions were ones the plan had already answered for itself:
    // it had seeded the Lists sheet with "Unassigned" placeholders exactly as
    // planner.prompt.ts instructs, then asked for the real values anyway. That
    // guidance says to raise such gaps in the FINAL SUMMARY, never to block on
    // them; enforcing it here means it no longer depends on the model
    // remembering.
    //
    // Safe because of TASKS.md #148: nothing reaches the workbook before
    // Accept. A plan the user can read and reject strictly dominates a dead
    // end — and the questions are still delivered, alongside the work rather
    // than instead of it.
    const openQuestions = plan.clarificationsNeeded.length > 0
      ? plan.clarificationsNeeded
      : plan.confidence === 'low'
        ? [
            plan.reasoning?.trim() ||
              'This request is ambiguous — what exactly should I change in the workbook?',
          ]
        : [];

    // Low confidence still blocks unconditionally — that is a separate, older
    // decision and the live failure gives no evidence against it (that plan's
    // confidence was "medium"). Low confidence means the model doubts its
    // READING of the request, where building 300 actions on a misreading wastes
    // minutes and hands back a plausible-looking wrong card. Unanswered
    // side-questions alongside a confident plan are a different thing entirely.
    const mustAsk =
      plan.confidence === 'low' ||
      (plan.clarificationsNeeded.length > 0 && plan.subtasks.length === 0);

    if (mustAsk) {
      emitter.send({ type: 'CLARIFY', questions: openQuestions });
      return {
        actions: [],
        iterationsRun: 0,
        verifierPassed: false,
        clarificationRequested: true,
        completedSubtasks: [],
        failedSubtask: null,
        partialProgress: false,
        planSubtasks: [],
        undeliveredSubtasks: [],
        openQuestions,
      };
    }

    emitter.send({
      type: 'CHECKPOINT',
      step: `${plan.subtasks.length} step${plan.subtasks.length > 1 ? 's' : ''} planned`,
    });

    const {
      actions: allActions,
      iterationsRun,
      verifierPassed,
      completedSubtasks,
      failedSubtask,
      partialProgress,
    } = await this.agenticLoop.run(prompt, plan.subtasks, context, emitter, {
      conversationId,
      correlationId: resolvedCorrelationId,
      toolEmit,
      usageTotals,
    });

    this.logger.log(
      `Agentic loop complete: ${allActions.length} actions, ${iterationsRun} iterations, verified: ${verifierPassed}, partial: ${partialProgress}`,
    );

    emitter.send({
      type: 'CHECKPOINT',
      step: `${allActions.length} actions ready for preview`,
    });

    const pruned = pruneSpuriousAddSheetActions(allActions);
    const clearAnnotated = annotateClearIntentOverwrite(pruned, prompt);
    const overwriteAnnotated = annotateExplicitOverwriteConfirmation(
      clearAnnotated,
      prompt,
      context.priorTurnActions ?? [],
    );

    return {
      actions: overwriteAnnotated,
      iterationsRun,
      verifierPassed,
      clarificationRequested: false,
      completedSubtasks,
      failedSubtask,
      partialProgress,
      // ONLY subtasks that actually produced actions. The Accept card renders
      // these as its promises (TASKS.md #149), so a subtask the Executor
      // silently emitted nothing for must not appear there — a live smoke test
      // caught the card promising "Write Consolidated Transactions header at
      // Main!A18" for a run whose actions never touched row 18. Describing the
      // PLAN rather than the DELIVERY is precisely the false-completeness shape
      // CODEBASE_ANALYSIS.md §3.7 keeps re-teaching. TASKS.md #155.
      planSubtasks: plan.subtasks
        .filter((s) =>
          completedSubtasks.some((c) => c.subtaskId === s.id && c.actions.length > 0),
        )
        .map((s) => ({
          id: s.id,
          description: s.description,
          targetSheet: s.targetSheet,
        })),
      // Questions raised but NOT blocked on: the build proceeded under a stated
      // assumption, and the user gets both. TASKS.md #171.
      openQuestions,
      /** Planned but delivered nothing — surfaced so the gap can be reported. */
      undeliveredSubtasks: plan.subtasks
        .filter(
          (s) =>
            !completedSubtasks.some((c) => c.subtaskId === s.id && c.actions.length > 0),
        )
        .map((s) => ({
          id: s.id,
          description: s.description,
          targetSheet: s.targetSheet,
        })),
    };
  }

  private resolveCorrelationId(value?: string): string {
    const trimmed = value?.trim();
    if (!trimmed || trimmed === '-') return `req_${Date.now()}`;
    return trimmed;
  }

  /**
   * Copies the accumulated usage onto the caller's telemetry out-param. A
   * no-op when the caller didn't pass one (e.g. existing callers/tests that
   * predate this wiring) — this is additive, never required.
   */
  private applyUsageToTelemetry(
    telemetry: LlmCallTelemetry | undefined,
    usageTotals: ReturnType<typeof createUsageAccumulator>,
  ): void {
    if (!telemetry) return;
    if (usageTotals.model) telemetry.model = usageTotals.model;
    if (usageTotals.promptTokens > 0 || usageTotals.completionTokens > 0) {
      telemetry.usage = {
        promptTokens: usageTotals.promptTokens,
        completionTokens: usageTotals.completionTokens,
        totalTokens: usageTotals.totalTokens,
      };
    }
  }
}
