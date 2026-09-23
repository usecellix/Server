import { Injectable, Logger, Optional } from '@nestjs/common';
import { PlannerAgent } from './planner.agent';
import { AgenticLoopService } from './agenticLoop.service';
import { SpecExtractorAgent } from './spec-extractor.agent';
import { SseEmitter } from './sse.emitter';
import { Action, AgentRunOptions, PlannerOutput } from './types/agent.types';
import { annotateExplicitOverwriteConfirmation } from '../excel-ai/utils/overwrite-confirmation.util';
import { annotateClearIntentOverwrite } from './utils/clear-intent-overwrite.util';
import { pruneSpuriousAddSheetActions } from './utils/compound-action.util';
import { createUsageAccumulator, UsageTotals } from './utils/usage-accumulator.util';
import { checkPlanIntegrity } from './utils/plan-integrity.util';
import { LlmCallTelemetry } from '../excel-ai/services/openrouter.service';

export interface OrchestratorRunResult {
  actions: Action[];
  iterationsRun: number;
  verifierPassed: boolean;
  clarificationRequested: boolean;
  completedSubtasks: Array<{ subtaskId: string; actions: Action[]; verified: boolean }>;
  failedSubtask: { subtaskId: string; reason: string } | null;
  /**
   * EVERY subtask that failed, not just the first — see
   * `AgenticLoopResult.failedSubtasks`'s docblock. TASKS.md #195. Note that
   * `undeliveredSubtasks` below already independently covers the FULL-SUCCESS
   * path's plan-vs-delivery gap; this field is what the PARTIAL-PROGRESS
   * (`!verifierPassed`) branch needs, since that branch never previously had
   * more than one failure's reason available at all.
   */
  failedSubtasks: Array<{ subtaskId: string; reason: string }>;
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
    @Optional() private readonly specExtractor?: SpecExtractorAgent,
  ) {}

  /**
   * Pins the user's own column lists onto the plan — LONG_PROMPT_RELIABILITY_PLAN.md
   * Phase 1.
   *
   * TASKS.md #281 — a turn that answers a clarifying question sends only the
   * short reply ("dd-mm-yyyy") as `prompt`; the ORIGINAL long request lives in
   * `conversationHistory` (TASKS.md #259 already saves it there for the
   * Planner's benefit). A live run showed the cost of feeding the raw short
   * reply here instead: `shouldExtractBuildSpec`'s length/subtask-count gate
   * saw a 10-character prompt and skipped Phase 1 entirely, and even if it
   * hadn't, extraction itself would have had no column list to find in
   * "dd-mm-yyyy" alone — so a resumed long build silently lost EVERY
   * reliability protection built this session (Phase 1's header pinning,
   * Phase 1.5's deterministic split) for that turn, and hit the same
   * near-timeout the original session's live runs did. Reconstructing the
   * turn's real prompt from prior user messages fixes both: the gate sees the
   * true length, and extraction has the actual column list to ground against.
   */
  private async withBuildSpec(
    prompt: string,
    plan: PlannerOutput,
    usageTotals: UsageTotals,
    conversationHistory: NonNullable<AgentRunOptions['conversationHistory']> = [],
  ): Promise<PlannerOutput> {
    if (!this.specExtractor) return plan;
    const priorUserText = conversationHistory
      .filter((entry) => entry.role === 'user')
      .map((entry) => entry.content)
      .join('\n');
    const effectivePrompt = priorUserText ? `${priorUserText}\n${prompt}` : prompt;
    return this.specExtractor.attach(effectivePrompt, plan, usageTotals);
  }

  /**
   * Phase 6 (TASKS.md #291) — a cheap, deterministic check that the plan is
   * not already wrong, run BEFORE any Executor call is spent on it. Every
   * other coverage net reasons from what the PLAN says; this one reasons from
   * what the USER asked for, which is the gap #285 fell through.
   */
  private applyPlanIntegrityGate(
    prompt: string,
    plan: PlannerOutput,
    context: AgentRunOptions['context'],
  ): PlannerOutput {
    const result = checkPlanIntegrity({ prompt, plan, context });

    for (const violation of result.violations) {
      const message = `Plan integrity (${violation.kind}): ${violation.detail}`;
      if (violation.fatal) this.logger.error(message);
      else this.logger.warn(message);
    }
    if (result.repaired.length > 0) {
      this.logger.log(
        `Plan integrity: recovered ${result.repaired.length} missing entit${
          result.repaired.length === 1 ? 'y' : 'ies'
        } before execution (${result.repaired.join(', ')}).`,
      );
    }
    if (result.violations.length === 0) {
      this.logger.log(`Plan integrity: ${plan.subtasks.length} step(s) passed pre-execution checks.`);
    }

    return result.plan;
  }

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
   * Plans without executing — the first half of a stepwise run
   * (STEPWISE_EXECUTION.md SD-1). Returns the plan plus the same
   * clarification/openQuestions handling `runDetailedWithUsage` applies, so the
   * stepwise path and the one-shot path cannot drift on when a request blocks
   * for a question versus proceeds under a stated assumption.
   */
  async planForStepwiseRun(
    opts: AgentRunOptions,
    emitter: SseEmitter,
    telemetry?: LlmCallTelemetry,
  ): Promise<{
    plan: PlannerOutput;
    openQuestions: string[];
    mustAsk: boolean;
  }> {
    const usageTotals = createUsageAccumulator();
    try {
      emitter.send({ type: 'THINKING', message: 'Planning your request...' });
      const plan = await this.withBuildSpec(
        opts.prompt,
        await this.planner.plan(
          opts.prompt,
          opts.context,
          opts.conversationHistory ?? [],
          opts.promptContext,
          this.resolveCorrelationId(opts.correlationId),
          opts.routerAssumption,
          opts.complexity,
          usageTotals,
          (summary) => emitter.send({ type: 'THINKING', message: summary }),
        ),
        usageTotals,
        opts.conversationHistory ?? [],
      );

      const gatedPlan = this.applyPlanIntegrityGate(opts.prompt, plan, opts.context);
      const openQuestions = this.resolveOpenQuestions(gatedPlan);
      const mustAsk = this.shouldBlockForClarification(gatedPlan);
      if (mustAsk) {
        emitter.send({ type: 'CLARIFY', questions: openQuestions });
      }
      return { plan: gatedPlan, openQuestions, mustAsk };
    } finally {
      this.applyUsageToTelemetry(telemetry, usageTotals);
    }
  }

  /**
   * Executes ONE wave of an already-planned stepwise run. Thin by design: the
   * verification machinery lives in AgenticLoopService.runWave and is shared
   * verbatim with the one-shot path, so this cannot become a second, subtly
   * different definition of "verified".
   */
  async runStepwiseWave(
    opts: AgentRunOptions & {
      waveSubtasks: PlannerOutput['subtasks'];
      priorActions: Array<{ subtask: PlannerOutput['subtasks'][number]; actions: Action[] }>;
    },
    emitter: SseEmitter,
    telemetry?: LlmCallTelemetry,
  ): Promise<{
    actions: Action[];
    completedSubtasks: Array<{ subtaskId: string; actions: Action[]; verified: boolean }>;
    failedSubtask: { subtaskId: string; reason: string } | null;
    /**
     * EVERY subtask this wave failed, not just the first — see
     * `AgenticLoopResult.failedSubtasks`'s docblock. TASKS.md #195: a wave of
     * many independent parallel subtasks (e.g. 12 month-sheet creates with no
     * dependsOn between them) can have several genuinely fail at once; using
     * only `failedSubtask` silently discarded every failure but one, with no
     * recorded reason anywhere for the rest.
     */
    failedSubtasks: Array<{ subtaskId: string; reason: string }>;
    verifierPassed: boolean;
  }> {
    const usageTotals = createUsageAccumulator();
    try {
      const result = await this.agenticLoop.runWave(
        opts.prompt,
        opts.waveSubtasks,
        opts.priorActions,
        opts.context,
        emitter,
        {
          conversationId: opts.conversationId,
          correlationId: this.resolveCorrelationId(opts.correlationId),
          toolEmit: opts.toolEmit,
          usageTotals,
          abortSignal: opts.abortSignal,
        },
      );

      const pruned = pruneSpuriousAddSheetActions(result.actions);
      const clearAnnotated = annotateClearIntentOverwrite(pruned, opts.prompt);
      const overwriteAnnotated = annotateExplicitOverwriteConfirmation(
        clearAnnotated,
        opts.prompt,
        opts.context.priorTurnActions ?? [],
      );

      return {
        actions: overwriteAnnotated,
        completedSubtasks: result.completedSubtasks,
        failedSubtask: result.failedSubtask,
        failedSubtasks: result.failedSubtasks,
        verifierPassed: result.verifierPassed,
      };
    } finally {
      this.applyUsageToTelemetry(telemetry, usageTotals);
    }
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
      onWaveComplete,
      precomputedPlan,
      abortSignal,
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
        onWaveComplete,
        precomputedPlan,
        abortSignal,
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
    /** Progressive per-wave emission — TASKS.md #174. */
    onWaveComplete: AgentRunOptions['onWaveComplete'],
    /** Already-planned output handed over by the stepwise gate — TASKS.md #196. */
    precomputedPlan?: PlannerOutput,
    abortSignal?: AbortSignal,
  ): Promise<OrchestratorRunResult> {
    // A plan already computed by the stepwise gate is reused rather than
    // re-derived. `tryStartStepwiseRun` plans, finds a single wave, declines,
    // and hands control here — which used to re-plan from scratch, costing a
    // second full Planner call (~1-2s and ~8k prompt tokens) on EVERY
    // single-wave Tier 3 request, i.e. essentially every simple one. The
    // original note called that "an accepted, bounded cost for keeping the two
    // paths from sharing mutable plan state"; the plan is handed over as a deep
    // copy instead, which removes the sharing without paying for the call.
    // TASKS.md #196.
    emitter.send({ type: 'THINKING', message: 'Planning your request...' });
    let plan: PlannerOutput;
    if (precomputedPlan) {
      plan = precomputedPlan;
    } else {
      plan = await this.withBuildSpec(
        prompt,
        await this.planner.plan(
          prompt,
          context,
          conversationHistory,
          promptContext,
          resolvedCorrelationId,
          routerAssumption,
          complexity,
          usageTotals,
          (summary) => emitter.send({ type: 'THINKING', message: summary }),
        ),
        usageTotals,
        conversationHistory,
      );
      plan = this.applyPlanIntegrityGate(prompt, plan, context);
    }

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
    const openQuestions = this.resolveOpenQuestions(plan);

    // Low confidence still blocks unconditionally — that is a separate, older
    // decision and the live failure gives no evidence against it (that plan's
    // confidence was "medium"). Low confidence means the model doubts its
    // READING of the request, where building 300 actions on a misreading wastes
    // minutes and hands back a plausible-looking wrong card. Unanswered
    // side-questions alongside a confident plan are a different thing entirely.
    const mustAsk = this.shouldBlockForClarification(plan);

    if (mustAsk) {
      emitter.send({ type: 'CLARIFY', questions: openQuestions });
      return {
        actions: [],
        iterationsRun: 0,
        verifierPassed: false,
        clarificationRequested: true,
        completedSubtasks: [],
        failedSubtask: null,
        failedSubtasks: [],
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
      failedSubtasks,
      partialProgress,
    } = await this.agenticLoop.run(prompt, plan.subtasks, context, emitter, {
      conversationId,
      correlationId: resolvedCorrelationId,
      toolEmit,
      usageTotals,
      onWaveComplete,
      abortSignal,
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
      failedSubtasks,
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
   * Questions raised but NOT blocked on — see the long rationale at the
   * `runDetailedWithUsage` call site (TASKS.md #171). Extracted so the stepwise
   * path applies the identical rule rather than a second copy of it.
   */
  private resolveOpenQuestions(plan: PlannerOutput): string[] {
    if (plan.clarificationsNeeded.length > 0) return plan.clarificationsNeeded;
    if (plan.confidence === 'low') {
      return [
        plan.reasoning?.trim() ||
          'This request is ambiguous — what exactly should I change in the workbook?',
      ];
    }
    return [];
  }

  /**
   * A large multi-step build (many sheets/tables/formulas) whose Planner still
   * has open questions — TASKS.md #259. Deliberately a NARROW addition on top
   * of #171's "don't block" rule, not a reversal of it: #171's regression was
   * a 23-subtask plan thrown away over 2 side-questions the plan had ALREADY
   * answered for itself with sensible defaults (placeholders, current year).
   * Losing that plan was the actual mistake — not the act of asking.
   *
   * Chosen by the user after that exact tradeoff was explained: for a request
   * this size (12 month sheets + dashboard + payment tracking is the live
   * case that prompted this), getting bank/unit names wrong across ~150
   * actions is expensive to redo, so asking once up front beats guessing and
   * hoping the summary note gets read. Small/simple edits (few subtasks) are
   * unaffected — they still never block, exactly as #171 intended.
   */
  private static readonly BIG_BUILD_SUBTASK_THRESHOLD = 6;

  private isBigAmbiguousBuild(plan: PlannerOutput): boolean {
    return (
      plan.subtasks.length >= OrchestratorService.BIG_BUILD_SUBTASK_THRESHOLD &&
      plan.clarificationsNeeded.length > 0
    );
  }

  /**
   * Blocks when there is nothing to build, the model doubts its reading, or
   * (TASKS.md #259) this is a big build with real open questions — see
   * `isBigAmbiguousBuild`'s docblock for why that last case is scoped
   * narrowly rather than reverting #171 wholesale.
   */
  private shouldBlockForClarification(plan: PlannerOutput): boolean {
    return (
      plan.confidence === 'low' ||
      (plan.clarificationsNeeded.length > 0 && plan.subtasks.length === 0) ||
      this.isBigAmbiguousBuild(plan)
    );
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
