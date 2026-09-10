import { Injectable, Logger, Optional } from '@nestjs/common';
import { PlannerFileLoggerService } from '../common/logging/planner-file-logger.service';
import { truncateForPlannerLog } from '../common/logging/planner-file-logger.util';
import { WorkflowTraceService } from '../common/logging/workflow-trace.service';
import { AppConfigService } from '../config/app-config.service';
import { PLANNER_RULES_ADDITION } from '../excel-ai/prompt/cellix-system-prompt';
import {
  LlmCompletionOutcome,
  OpenRouterService,
} from '../excel-ai/services/openrouter.service';
import {
  PLANNER_COST_CAP_USER_MESSAGE,
  PLANNER_EXHAUSTED_USER_MESSAGE,
  PlannerCostCapExceededError,
  PlannerExhaustedError,
} from './errors';
import {
  PLANNER_RULES_BODY,
  PLANNER_SYSTEM_PROMPT,
  buildPlannerUserMessage,
} from './prompts/planner.prompt';
import {
  PLANNER_COARSE_SYSTEM_PROMPT,
  buildCoarsePlannerUserMessage,
  buildPhaseExpansionSystemPrompt,
  buildPhaseExpansionUserMessage,
} from './prompts/planner-phase.prompt';
import { parseAgentJson } from './utils/parse-agent-json.util';
import { buildCompoundFallbackSubtasks } from './utils/compound-action.util';
import { ensureNumberFormatPlanSafety } from './utils/preserve-number-format.util';
import { ensureReferencedSheetsPlanned, ensureRepeatForCoverage } from './utils/plan-coverage.util';
import {
  PLANNER_LAST_RESORT_MAX_TOKENS,
  PLANNER_REASONING_MAX_TOKENS,
  resolvePlannerMaxTokens,
} from './utils/planner-token-budget.util';
import {
  COST_CAP_USD,
  estimateLlmCallCostUsd,
  resolvePricingForModel,
} from '../excel-ai/llm/model-router';
import { addUsage, UsageTotals } from './utils/usage-accumulator.util';
import { CoarsePlanOutput, PlanPhase, PlannerOutput, SubTask, WorkbookContext } from './types/agent.types';
import { resolveTier3ComplexityScore } from './utils/planner-token-budget.util';
import { StructuredLogger } from './logging/structured-logger';

const JSON_RETRY_SUFFIX =
  '\n\nIMPORTANT: Your previous response was not valid JSON. Reply with ONLY a single JSON object matching the schema — no markdown fences, no commentary.';

/**
 * Short, incremental progress callback for two-pass planning (TASKS.md #193).
 * A large build's planning phase is now several sequential LLM calls (one
 * coarse pass + one per phase) that can each take real wall-clock time — with
 * nothing surfaced to the user in between, the ONLY signal for the whole
 * planning phase was a single static "Planning your request..." message sent
 * once at the very start by `OrchestratorService`. This callback lets the
 * Planner report each small step as it completes (not big chunks of
 * reasoning text — one short phrase per phase) so a caller wired to an SSE
 * emitter (see `OrchestratorService.planForStepwiseRun`/`runDetailedWithUsage`)
 * can turn planning into a short, legible progress trail instead of a single
 * multi-minute silence. Optional and additive — omitting it changes nothing.
 */
export type PlannerProgressCallback = (summary: string) => void;

/**
 * Two-pass planning (TASKS.md #191) trigger — `resolveTier3ComplexityScore`
 * ALONE is not enough: it conflates "many distinct sections on one sheet"
 * (single-pass handles this fine, e.g. a dashboard with a summary+analysis+
 * pivot+chart, score 4-6) with "one structure repeated many times" (the
 * actual repeated-truncation shape, TASKS.md #170/#187/#190 — 12 month
 * sheets + a multi-section dashboard, ~19 subtasks). Both
 * `spec16-fallback-unreachable.spec.ts` benchmark prompts that also score 6
 * ("add a sheet for each month with a ledger table...", "build a purchase
 * register dashboard: a summary, an analysis section...") are real, already-
 * well-tested prompts that fit comfortably in ONE single-pass plan (4-8
 * subtasks) — scoring high on object-count is not the same as needing
 * decomposition.
 *
 * The length signal is what actually separates them: the real incident's
 * prompt is 369 characters (it spells out 10 column names verbatim) against
 * a maximum of 105 characters across every existing benchmark prompt that
 * also scores >= 4. A long, compound prompt correlates with many long
 * subtask descriptions in the resulting plan — the actual thing that
 * exhausts the completion budget — far better than object-count alone.
 * Requiring BOTH keeps two-pass planning off ordinary compound requests and
 * reserves it for requests that are both compound AND verbose enough to risk
 * truncation, rather than reopening this exact regression at a different
 * score cutoff.
 */
const TWO_PASS_COMPLEXITY_SCORE_THRESHOLD = 4;
const TWO_PASS_PROMPT_LENGTH_THRESHOLD = 200;

function needsTwoPassPlanning(prompt: string): boolean {
  return (
    resolveTier3ComplexityScore(prompt) >= TWO_PASS_COMPLEXITY_SCORE_THRESHOLD &&
    prompt.length >= TWO_PASS_PROMPT_LENGTH_THRESHOLD
  );
}

/** Coarse-pass budget — a phase list is 5-8 short entries, nowhere near this. */
const COARSE_PLAN_MAX_TOKENS = 2048;

/** Per-phase expansion budget — one phase's worth of subtasks, not the whole plan. */
const PHASE_EXPANSION_MAX_TOKENS = 4096;

@Injectable()
export class PlannerAgent {
  private readonly logger = new Logger(PlannerAgent.name);

  constructor(
    private readonly llm: OpenRouterService,
    private readonly config: AppConfigService,
    private readonly structuredLogger: StructuredLogger = new StructuredLogger(),
    @Optional() private readonly plannerFileLogger?: PlannerFileLoggerService,
    @Optional() private readonly workflowTrace?: WorkflowTraceService,
  ) {}

  async plan(
    prompt: string,
    context: WorkbookContext,
    history: { role: string; content: string }[] = [],
    promptContext?: string,
    correlationId = `req_${Date.now()}`,
    routerAssumption?: string,
    complexity?: 0 | 1 | 2 | 3,
    /** Out-param, same pattern as `LlmCompletionOutcome` — accumulates real
     * promptTokens/completionTokens across every attempt this call makes, so
     * a caller (OrchestratorService) can report actual usage instead of the
     * 0 every Tier-3 audit_logs row previously carried. */
    usageTotals?: UsageTotals,
    /** Optional short-progress callback — see `PlannerProgressCallback`'s own
     *  docblock. Only consulted by the two-pass path; a single-pass plan is
     *  one call and has nothing incremental to report. */
    onProgress?: PlannerProgressCallback,
  ): Promise<PlannerOutput> {
    // Two-pass planning (TASKS.md #191) — a single-pass plan for a large
    // compound build (12 month sheets + a multi-section dashboard) has to
    // describe everything in one JSON response and repeatedly truncates even
    // at the last-resort ceiling (TASKS.md #170/#187/#190's incidents are all
    // the same root cause). See `needsTwoPassPlanning`'s own docblock for why
    // this needs both a compound-object signal AND a length signal, not
    // object-count alone. Only Tier 3 (or unclassified, which defaults to
    // Tier 3 sizing) is ever large enough to need this; Tier 0-2 requests are
    // structurally too small to trigger either signal.
    if ((complexity === undefined || complexity === 3) && needsTwoPassPlanning(prompt)) {
      return this.planTwoPass(
        prompt,
        context,
        history,
        promptContext,
        correlationId,
        routerAssumption,
        usageTotals,
        onProgress,
      );
    }

    const startedAt = Date.now();
    // Spec 16 fix #2: Planner calls resolve to their own model override
    // (`OPENROUTER_MODEL_PLANNER`) rather than unconditionally sharing
    // `openRouterModelHigh` with the Executor. Defaults to `openRouterModelHigh`
    // (openai/gpt-5) unchanged — this exists so a non-reasoning/lighter-reasoning
    // model can be evaluated for the Planner's specific job (structured JSON
    // decomposition, not open-ended reasoning) without touching Executor
    // behavior. See `AppConfigService.openRouterModelPlanner`.
    const model = this.config.openRouterModelPlanner;
    const systemPrompt = PLANNER_SYSTEM_PROMPT + PLANNER_RULES_ADDITION;
    let userMessage = buildPlannerUserMessage(prompt, context, history, promptContext);
    if (routerAssumption) {
      userMessage = `[Router assumption: ${routerAssumption}]\n\n${userMessage}`;
    }

    const maxTokens = resolvePlannerMaxTokens(complexity, prompt);

    // Spec 16 fix #4: refuse the call up front if it would exceed the shared
    // per-call cost cap, rather than never checking at all (PlannerAgent
    // previously never consulted ModelRouter/COST_CAP_USD). Priced against the
    // WORST CASE the retry ladder below could reach — PLANNER_LAST_RESORT_MAX_TOKENS,
    // not this call's possibly-smaller `maxTokens` — because promptTokens is
    // fixed across every retry (same message resent) while completionTokens
    // only grows; checking once against the ceiling avoids refusing only after
    // the user has already waited through the cheaper attempts. Pricing is
    // resolved from the model actually configured for the Planner
    // (`openRouterModelPlanner`), not a hardcoded HIGH-tier assumption — that
    // model can diverge from `openRouterModelHigh` via fix #2's eval override.
    const promptTokenEstimate = Math.ceil((systemPrompt.length + userMessage.length) / 4);
    const { pricing, approximate } = resolvePricingForModel(model, this.config);
    if (approximate) {
      this.logger.warn(
        `Planner cost estimate is APPROXIMATE — model=${model} does not match ` +
          `openRouterModelLow/Medium/High; using HIGH pricing as a conservative default.`,
      );
    }
    const worstCaseCostUsd = estimateLlmCallCostUsd(
      pricing,
      promptTokenEstimate,
      PLANNER_LAST_RESORT_MAX_TOKENS,
    );
    if (worstCaseCostUsd > COST_CAP_USD) {
      this.logger.error(
        `Planner refusing call — estimated worst-case cost $${worstCaseCostUsd.toFixed(4)} ` +
          `exceeds cap $${COST_CAP_USD} (model=${model}, promptTokens~${promptTokenEstimate}, ` +
          `worstCaseCompletionTokens=${PLANNER_LAST_RESORT_MAX_TOKENS}${approximate ? ', pricing approximate' : ''})`,
      );
      throw new PlannerCostCapExceededError(PLANNER_COST_CAP_USER_MESSAGE, {
        originalMessage: prompt,
        estimatedCostUsd: worstCaseCostUsd,
        costCapUsd: COST_CAP_USD,
        promptTokens: promptTokenEstimate,
        maxTokens: PLANNER_LAST_RESORT_MAX_TOKENS,
      });
    }

    const completeOpts = {
      systemPrompt,
      model,
      maxTokens,
      // Planner stays at full reasoning capability — do not lower this to save
      // tokens (per explicit instruction, distinct from Executor/Verifier/Tier1
      // below, which are deliberately capped at low/none). reasoningMaxTokens
      // still bounds spend regardless of effort level, so this raises reasoning
      // depth within the same budget ceiling rather than uncapping it.
      reasoningEffort: 'high' as const,
      reasoningMaxTokens: PLANNER_REASONING_MAX_TOKENS,
    };

    // A plan cut off by the token budget frequently still parses — the model
    // closes the JSON it has emitted so far — so parse success alone is NOT
    // evidence the plan is complete. Treat truncation as a failure to be
    // retried on a bigger budget, exactly like a parse failure.
    if (usageTotals) usageTotals.model ??= model;

    const outcome: LlmCompletionOutcome = {};
    let raw = await this.llm.complete({
      ...completeOpts,
      userMessage,
      temperature: 0.2,
      outcome,
    });
    addUsage(usageTotals, outcome.usage);
    this.structuredLogger.debugRawResponse(correlationId, 'planner', model, raw);

    let retried = false;
    let lastResort = false;
    let truncated = outcome.truncated === true;
    let parsed = truncated ? null : this.tryParsePlanner(raw, correlationId, model);
    if (!parsed) {
      retried = true;
      this.logger.warn(
        truncated
          ? `Planner output truncated at maxTokens=${maxTokens} (parseable but incomplete) — retrying once.`
          : `Planner JSON parse failed — retrying once. Raw snippet: ${this.clip(raw)}`,
      );
      const retryOutcome: LlmCompletionOutcome = {};
      raw = await this.llm.complete({
        ...completeOpts,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        temperature: 0.1,
        outcome: retryOutcome,
      });
      addUsage(usageTotals, retryOutcome.usage);
      this.structuredLogger.debugRawResponse(correlationId, 'planner', model, raw);
      truncated = retryOutcome.truncated === true;
      parsed = truncated ? null : this.tryParsePlanner(raw, correlationId, model);
    }

    if (!parsed) {
      lastResort = true;
      this.logger.warn(
        `Planner still empty/unparseable — last-resort retry with maxTokens=${PLANNER_LAST_RESORT_MAX_TOKENS}`,
      );
      const lastResortOutcome: LlmCompletionOutcome = {};
      raw = await this.llm.complete({
        ...completeOpts,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        temperature: 0.1,
        maxTokens: PLANNER_LAST_RESORT_MAX_TOKENS,
        reasoningMaxTokens: Math.min(PLANNER_REASONING_MAX_TOKENS, 768),
        outcome: lastResortOutcome,
      });
      addUsage(usageTotals, lastResortOutcome.usage);
      this.structuredLogger.debugRawResponse(correlationId, 'planner', model, raw);
      truncated = lastResortOutcome.truncated === true;
      // Last resort: a truncated plan here is still better than no plan, so we
      // keep it rather than failing the request outright — but it must not pass
      // as complete. Flag it so the caller/log records that this plan may be
      // short, instead of the silent under-planning this whole check exists for.
      parsed = this.tryParsePlanner(raw, correlationId, model);
      if (parsed && truncated) {
        this.logger.error(
          `Planner output STILL truncated at last-resort maxTokens=${PLANNER_LAST_RESORT_MAX_TOKENS} — ` +
            `proceeding with a possibly incomplete plan (${parsed.subtasks.length} subtasks). ` +
            `The request may be under-planned; consider splitting it.`,
        );
      }
    }

    if (parsed) {
      this.logger.log(
        `Planner produced ${parsed.subtasks.length} subtasks, confidence: ${parsed.confidence}`,
      );
      const covered = this.pruneUnsatisfiableSubtasks(
        this.planReferencedSheets(
          ensureNumberFormatPlanSafety(prompt, this.ensureMultiClauseCoverage(prompt, parsed)),
          context,
        ),
      );
      this.structuredLogger.logAgentEvent({
        correlationId,
        agent: 'planner',
        model,
        durationMs: Date.now() - startedAt,
        success: true,
        tokenUsage: this.structuredLogger.estimateTokens(raw),
        rawResponse: raw,
        parsedResponse: covered,
      });
      this.recordWorkflowNode(correlationId, startedAt, true, prompt, context, covered, model);
      this.writePlannerFileLog({
        correlationId,
        model,
        durationMs: Date.now() - startedAt,
        success: true,
        prompt,
        context,
        history,
        promptContext,
        routerAssumption,
        userMessage,
        systemPrompt,
        raw,
        parsed: covered,
        fallback: false,
        retried: retried || lastResort,
      });
      return covered;
    }

    // Prefer useful structured fallbacks (empty sheet clarify / create+sort) over a stub.
    const usefulFallback = this.tryUsefulFallbackPlan(prompt, context);
    if (usefulFallback) {
      this.logger.warn(
        `Planner JSON failed after retries — using structured fallback (${usefulFallback.reasoning}). Raw snippet: ${this.clip(raw)}`,
      );
      this.structuredLogger.logAgentEvent({
        correlationId,
        agent: 'planner',
        model,
        durationMs: Date.now() - startedAt,
        success: false,
        tokenUsage: this.structuredLogger.estimateTokens(raw),
        rawResponse: raw,
        parsedResponse: usefulFallback,
        error: 'Planner JSON parse failed after retry — structured fallback',
      });
      this.recordWorkflowNode(
        correlationId,
        startedAt,
        false,
        prompt,
        context,
        usefulFallback,
        model,
        'Planner JSON parse failed after retry — structured fallback',
      );
      this.writePlannerFileLog({
        correlationId,
        model,
        durationMs: Date.now() - startedAt,
        success: false,
        error: 'Planner JSON parse failed after retry — structured fallback',
        prompt,
        context,
        history,
        promptContext,
        routerAssumption,
        userMessage,
        systemPrompt,
        raw,
        parsed: usefulFallback,
        fallback: true,
        retried: true,
      });
      return usefulFallback;
    }

    this.logger.error(
      `Planner exhausted after retries — refusing stub plan. Raw snippet: ${this.clip(raw)}`,
    );
    this.structuredLogger.logAgentEvent({
      correlationId,
      agent: 'planner',
      model,
      durationMs: Date.now() - startedAt,
      success: false,
      tokenUsage: this.structuredLogger.estimateTokens(raw),
      rawResponse: raw,
      error: 'PlannerExhaustedError',
    });
    this.recordWorkflowNode(
      correlationId,
      startedAt,
      false,
      prompt,
      context,
      { subtasks: [], clarificationsNeeded: [], confidence: 'low', reasoning: '' },
      model,
      'PlannerExhaustedError',
    );
    this.writePlannerFileLog({
      correlationId,
      model,
      durationMs: Date.now() - startedAt,
      success: false,
      error: 'PlannerExhaustedError',
      prompt,
      context,
      history,
      promptContext,
      routerAssumption,
      userMessage,
      systemPrompt,
      raw,
      parsed: {
        subtasks: [],
        clarificationsNeeded: [],
        confidence: 'low',
        reasoning: 'PlannerExhaustedError',
      },
      fallback: false,
      retried: true,
    });

    throw new PlannerExhaustedError(PLANNER_EXHAUSTED_USER_MESSAGE, {
      originalMessage: prompt,
    });
  }

  /**
   * Two-pass planning (TASKS.md #191). Pass 1 identifies coarse phases (a
   * handful of short entries — see `planner-phase.prompt.ts`'s own docblock
   * for why this call is small and reliable almost by construction). Pass 2
   * expands each phase into real subtasks, ONE PHASE AT A TIME, in dependency
   * order — each expansion call is scoped to a single phase's worth of
   * output, so no single call has to describe 19 things at once the way a
   * single-pass plan for this request shape does.
   *
   * Phases are expanded SEQUENTIALLY, not in parallel waves: a phase that
   * depends on another needs that phase's REAL subtask ids (not just its
   * phase id) to write correct `dependsOn` references, so the dependency
   * order has to be resolved before expansion, not after. This trades wall-
   * clock time (N small sequential calls instead of 1-3 large ones) for
   * reliability — an accepted, explicit tradeoff (STEPWISE_EXECUTION.md's
   * SD-3 makes the identical trade for the same reason).
   *
   * Returns a NORMAL `PlannerOutput` (flat `subtasks[]`) — nothing downstream
   * (`computeExecutionWaves`, `pruneUnsatisfiableSubtasks`, the Orchestrator,
   * stepwise execution) needs to know two-pass planning happened at all.
   */
  private async planTwoPass(
    prompt: string,
    context: WorkbookContext,
    history: { role: string; content: string }[],
    promptContext: string | undefined,
    correlationId: string,
    routerAssumption: string | undefined,
    usageTotals: UsageTotals | undefined,
    onProgress: PlannerProgressCallback | undefined,
  ): Promise<PlannerOutput> {
    const startedAt = Date.now();
    const model = this.config.openRouterModelPlanner;
    if (usageTotals) usageTotals.model ??= model;

    onProgress?.('Sketching the overall plan…');
    const coarse = await this.planCoarse(
      prompt,
      context,
      history,
      promptContext,
      correlationId,
      routerAssumption,
      model,
      usageTotals,
    );

    if (coarse.clarificationsNeeded.length > 0 && coarse.phases.length === 0) {
      this.logger.log(
        `Two-pass planning: coarse pass raised a blocking clarification, no phases — returning as-is.`,
      );
      return {
        subtasks: [],
        clarificationsNeeded: coarse.clarificationsNeeded,
        confidence: coarse.confidence,
        reasoning: coarse.reasoning,
      };
    }

    const mergedPhases = this.mergeSameSheetPhases(coarse.phases);

    this.logger.log(
      `Two-pass planning: coarse pass identified ${coarse.phases.length} phase(s) ` +
        `(${mergedPhases.length} after same-sheet merge) — expanding each.`,
    );
    onProgress?.(
      mergedPhases.length === 1
        ? 'Planned 1 step — working out the details…'
        : `Planned ${mergedPhases.length} steps — working out the details…`,
    );

    const ordered = this.orderPhasesByDependency(mergedPhases);
    const subtasksByPhase = new Map<string, SubTask[]>();
    const allSubtasks: SubTask[] = [];
    const clarifications: string[] = [...coarse.clarificationsNeeded];
    let phaseIndex = 0;

    for (const phase of ordered) {
      phaseIndex += 1;
      onProgress?.(
        `Working out step ${phaseIndex}/${ordered.length}: ${this.summarizePhaseForProgress(phase)}…`,
      );

      const dependencySubtaskIds = phase.dependsOn.flatMap(
        (depPhaseId) => subtasksByPhase.get(depPhaseId) ?? [],
      ).map((s) => s.id);

      const expansion = await this.expandPhase(
        prompt,
        context,
        history,
        promptContext,
        correlationId,
        phase,
        dependencySubtaskIds,
        model,
        usageTotals,
      );

      // TASKS.md #229 — a repeatFor phase must cover EVERY entry; the expansion
      // prompt asking for it is not enough (live: 12 months planned, only
      // January expanded). Done before stitching so dependent phases see all
      // entries' real subtask ids.
      const repeatCoverage = ensureRepeatForCoverage(phase, expansion.subtasks);
      if (repeatCoverage.filled.length > 0) {
        this.logger.warn(
          `Two-pass planning: phase "${phase.id}" expanded only part of its repeatFor — cloned ` +
            `"${repeatCoverage.template}" for ${repeatCoverage.filled.length} missing entr` +
            `${repeatCoverage.filled.length === 1 ? 'y' : 'ies'}: ${repeatCoverage.filled.join(', ')}.`,
        );
      }

      const namespaced = this.namespacePhaseSubtasks(phase.id, repeatCoverage.subtasks);
      // Any subtask this phase produced with NO local dependsOn is wired to
      // depend on the phases IT depends on, mirroring what a single-pass plan
      // writes by hand (e.g. Main's totals subtask depending on all 12 month-
      // create subtasks) — automated here since the model only sees THIS
      // phase's own output, not the full cross-phase picture.
      const stitched =
        dependencySubtaskIds.length > 0
          ? namespaced.map((s) =>
              s.dependsOn.length === 0 ? { ...s, dependsOn: dependencySubtaskIds } : s,
            )
          : namespaced;

      subtasksByPhase.set(phase.id, stitched);
      allSubtasks.push(...stitched);
      clarifications.push(...expansion.clarificationsNeeded);
    }

    const merged: PlannerOutput = {
      subtasks: allSubtasks,
      clarificationsNeeded: clarifications,
      confidence: coarse.confidence,
      reasoning: coarse.reasoning,
    };

    const covered = this.pruneUnsatisfiableSubtasks(
      this.planReferencedSheets(
        ensureNumberFormatPlanSafety(prompt, this.ensureMultiClauseCoverage(prompt, merged)),
        context,
      ),
    );

    onProgress?.('Plan ready — starting the build…');

    this.logger.log(
      `Two-pass planning complete: ${covered.subtasks.length} subtask(s) across ${ordered.length} ` +
        `phase(s), durationMs=${Date.now() - startedAt}`,
    );
    this.recordWorkflowNode(correlationId, startedAt, true, prompt, context, covered, model);
    this.writePlannerFileLog({
      correlationId,
      model,
      durationMs: Date.now() - startedAt,
      success: true,
      prompt,
      context,
      history,
      promptContext,
      routerAssumption,
      userMessage: `[two-pass: ${ordered.length} phases]`,
      systemPrompt: PLANNER_COARSE_SYSTEM_PROMPT,
      raw: JSON.stringify(coarse),
      parsed: covered,
      fallback: false,
      retried: false,
    });

    return covered;
  }

  /** Pass 1 — coarse phase identification. Small budget, minimal retry ladder
   *  (a phase list is short enough that truncation here is rare). */
  private async planCoarse(
    prompt: string,
    context: WorkbookContext,
    history: { role: string; content: string }[],
    promptContext: string | undefined,
    correlationId: string,
    routerAssumption: string | undefined,
    model: string,
    usageTotals: UsageTotals | undefined,
  ): Promise<CoarsePlanOutput> {
    let userMessage = buildCoarsePlannerUserMessage(prompt, context, history, promptContext);
    if (routerAssumption) {
      userMessage = `[Router assumption: ${routerAssumption}]\n\n${userMessage}`;
    }

    const outcome: LlmCompletionOutcome = {};
    let raw = await this.llm.complete({
      systemPrompt: PLANNER_COARSE_SYSTEM_PROMPT,
      model,
      maxTokens: COARSE_PLAN_MAX_TOKENS,
      reasoningEffort: 'high',
      reasoningMaxTokens: PLANNER_REASONING_MAX_TOKENS,
      userMessage,
      temperature: 0.2,
      outcome,
    });
    addUsage(usageTotals, outcome.usage);

    let parsed = outcome.truncated ? null : this.tryParseCoarsePlan(raw, correlationId, model);
    if (!parsed) {
      this.logger.warn(
        `Two-pass planning: coarse pass ${outcome.truncated ? 'truncated' : 'unparseable'} — retrying once.`,
      );
      const retryOutcome: LlmCompletionOutcome = {};
      raw = await this.llm.complete({
        systemPrompt: PLANNER_COARSE_SYSTEM_PROMPT,
        model,
        maxTokens: COARSE_PLAN_MAX_TOKENS * 2,
        reasoningEffort: 'high',
        reasoningMaxTokens: PLANNER_REASONING_MAX_TOKENS,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        temperature: 0.1,
        outcome: retryOutcome,
      });
      addUsage(usageTotals, retryOutcome.usage);
      parsed = retryOutcome.truncated ? null : this.tryParseCoarsePlan(raw, correlationId, model);
    }

    if (parsed) return parsed;

    // A coarse pass that fails entirely is rare (the whole point is that it's
    // small) — fall back to treating the request as one single phase, which
    // routes it back through ordinary single-pass planning for THIS phase's
    // expansion call. Not a full single-pass plan (that already failed
    // upstream, which is WHY two-pass triggered) — one phase covering
    // everything, expanded with the full detailed rules.
    this.logger.error(
      `Two-pass planning: coarse pass failed after retry — falling back to a single all-covering phase.`,
    );
    return {
      phases: [{ id: 'p1', kind: prompt, targetSheet: context.activeSheetName || 'Sheet1', dependsOn: [] }],
      clarificationsNeeded: [],
      confidence: 'medium',
      reasoning: 'Coarse planning pass failed — treating the whole request as one phase.',
    };
  }

  /** Pass 2 — expand ONE phase into real subtasks, reusing the full rule body. */
  private async expandPhase(
    originalPrompt: string,
    context: WorkbookContext,
    history: { role: string; content: string }[],
    promptContext: string | undefined,
    correlationId: string,
    phase: PlanPhase,
    dependencySubtaskIds: string[],
    model: string,
    usageTotals: UsageTotals | undefined,
  ): Promise<PlannerOutput> {
    const systemPrompt = buildPhaseExpansionSystemPrompt(PLANNER_RULES_BODY + PLANNER_RULES_ADDITION);
    const userMessage = buildPhaseExpansionUserMessage({
      originalPrompt,
      context,
      history,
      promptContext,
      phase,
      dependencySubtaskIds,
    });

    const outcome: LlmCompletionOutcome = {};
    let raw = await this.llm.complete({
      systemPrompt,
      model,
      maxTokens: PHASE_EXPANSION_MAX_TOKENS,
      reasoningEffort: 'high',
      reasoningMaxTokens: PLANNER_REASONING_MAX_TOKENS,
      userMessage,
      temperature: 0.2,
      outcome,
    });
    addUsage(usageTotals, outcome.usage);

    let parsed = outcome.truncated ? null : this.tryParsePlanner(raw, correlationId, model);
    if (!parsed) {
      this.logger.warn(
        `Two-pass planning: phase "${phase.id}" (${phase.kind}) ` +
          `${outcome.truncated ? 'truncated' : 'unparseable'} — retrying once at a larger budget.`,
      );
      const retryOutcome: LlmCompletionOutcome = {};
      raw = await this.llm.complete({
        systemPrompt,
        model,
        maxTokens: Math.min(PHASE_EXPANSION_MAX_TOKENS * 2, PLANNER_LAST_RESORT_MAX_TOKENS),
        reasoningEffort: 'high',
        reasoningMaxTokens: PLANNER_REASONING_MAX_TOKENS,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        temperature: 0.1,
        outcome: retryOutcome,
      });
      addUsage(usageTotals, retryOutcome.usage);
      parsed = this.tryParsePlanner(raw, correlationId, model);
      if (parsed && retryOutcome.truncated) {
        this.logger.error(
          `Two-pass planning: phase "${phase.id}" STILL truncated after retry — ` +
            `proceeding with ${parsed.subtasks.length} subtask(s) from this phase; may be incomplete.`,
        );
      }
    }

    if (parsed) return parsed;

    this.logger.error(
      `Two-pass planning: phase "${phase.id}" (${phase.kind}) produced nothing usable after retry — skipping this phase.`,
    );
    return {
      subtasks: [],
      clarificationsNeeded: [
        `I could not plan the "${phase.kind}" part of this request — the response was cut off or malformed. Ask again to add it once you see what's here.`,
      ],
      confidence: 'low',
      reasoning: `Phase "${phase.id}" expansion failed.`,
    };
  }

  /** Namespaces a phase's locally-unique subtask ids (s1, s2, ...) to avoid
   *  collisions once merged with every other phase's output, rewriting each
   *  subtask's own internal dependsOn references to match. */
  private namespacePhaseSubtasks(phaseId: string, subtasks: SubTask[]): SubTask[] {
    const idMap = new Map(subtasks.map((s) => [s.id, `${phaseId}_${s.id}`]));
    return subtasks.map((s) => ({
      ...s,
      id: idMap.get(s.id)!,
      dependsOn: s.dependsOn.map((dep) => idMap.get(dep) ?? dep),
    }));
  }

  /**
   * Topological order for phases by their OWN dependsOn edges — sequential,
   * not waved, since a later phase's expansion call needs an EARLIER phase's
   * REAL subtask ids (see `planTwoPass`'s own docblock for why parallel waves
   * don't work here the way `computeExecutionWaves` does for execution).
   * A cycle or dangling reference just leaves the offending phase at the end
   * rather than throwing — better to expand it (with an empty dependency
   * list, if its real dependency never resolved) than to fail the whole plan.
   */
  private orderPhasesByDependency(phases: PlanPhase[]): PlanPhase[] {
    const byId = new Map(phases.map((p) => [p.id, p]));
    const visited = new Set<string>();
    const ordered: PlanPhase[] = [];

    const visit = (id: string, stack: Set<string>) => {
      if (visited.has(id) || stack.has(id)) return;
      const phase = byId.get(id);
      if (!phase) return;
      stack.add(id);
      for (const dep of phase.dependsOn) {
        visit(dep, stack);
      }
      stack.delete(id);
      visited.add(id);
      ordered.push(phase);
    };

    for (const phase of phases) {
      visit(phase.id, new Set());
    }

    return ordered;
  }

  /**
   * Structural safety net for the "ONE SHEET'S BUILD IS ONE PHASE" rule in
   * `PLANNER_COARSE_SYSTEM_PROMPT` — that rule is prompt-only and therefore
   * probabilistic; this makes the invariant hold even when the coarse pass
   * ignores it. Observed in production: a coarse pass split one Main-sheet
   * build into "consolidate details into Main" + "build dashboard/KPIs on
   * Main" as two separate phases. Each phase's expansion call cannot see what
   * the OTHER phase's expansion actually wrote, so both independently
   * decided they needed the same totals table and both wrote it — the second
   * phase's writes were then rejected by the overwrite guard as duplicates.
   *
   * Phases that share a `targetSheet` (and are NOT `repeatFor` phases — those
   * intentionally reuse one representative sheet name to describe work that
   * actually spans many DIFFERENT real sheets, e.g. "January" standing in for
   * all 12 months) are merged into one phase before any expansion call runs,
   * so a single expansion call sees the whole sheet's build at once and can't
   * duplicate work against itself.
   */
  private mergeSameSheetPhases(phases: PlanPhase[]): PlanPhase[] {
    const bySheet = new Map<string, PlanPhase[]>();
    for (const phase of phases) {
      if (phase.repeatFor?.length) continue;
      const group = bySheet.get(phase.targetSheet) ?? [];
      group.push(phase);
      bySheet.set(phase.targetSheet, group);
    }

    const idRemap = new Map<string, string>();
    const mergedById = new Map<string, PlanPhase>();
    const droppedIds = new Set<string>();

    for (const group of bySheet.values()) {
      if (group.length === 1) continue;

      const survivor = group[0];
      for (const dropped of group.slice(1)) {
        idRemap.set(dropped.id, survivor.id);
        droppedIds.add(dropped.id);
      }

      const dependsOn = Array.from(
        new Set(group.flatMap((p) => p.dependsOn).filter((dep) => !group.some((p) => p.id === dep))),
      );

      this.logger.warn(
        `Two-pass planning: merging ${group.length} phases that all target sheet "${survivor.targetSheet}" ` +
          `(${group.map((p) => p.id).join(', ')}) into one phase "${survivor.id}" to prevent duplicate builds.`,
      );

      mergedById.set(survivor.id, {
        id: survivor.id,
        kind: group.map((p) => p.kind).join(' Also: '),
        targetSheet: survivor.targetSheet,
        dependsOn,
      });
    }

    // Preserve the coarse pass's original phase order — only substitute the
    // merged phase where the survivor was, and drop the phases it absorbed.
    return phases
      .filter((phase) => !droppedIds.has(phase.id))
      .map((phase) => mergedById.get(phase.id) ?? phase)
      .map((phase) => ({
        ...phase,
        dependsOn: Array.from(
          new Set(phase.dependsOn.map((dep) => idRemap.get(dep) ?? dep).filter((dep) => dep !== phase.id)),
        ),
      }));
  }

  /**
   * A phase's own `kind` text is written for the EXPANSION call's prompt, not
   * for a person watching progress — it can run to a full sentence ("Create
   * the Main sheet that consolidates all details from the 12 month sheets
   * into one combined view..."). This trims it to something short enough for
   * a one-line status update, per the user's ask for small, summarized
   * progress text rather than big chunks during planning.
   */
  private summarizePhaseForProgress(phase: PlanPhase): string {
    const label = phase.repeatFor?.length
      ? `${phase.targetSheet} (+${phase.repeatFor.length - 1} more)`
      : phase.targetSheet;
    const firstClause = phase.kind.split(/[.;]/)[0].trim();
    const short = firstClause.length <= 60 ? firstClause : `${firstClause.slice(0, 57)}...`;
    return `${label} — ${short}`;
  }

  private tryParseCoarsePlan(
    raw: string,
    correlationId: string,
    model: string,
  ): CoarsePlanOutput | null {
    try {
      const parsed = parseAgentJson<Partial<CoarsePlanOutput>>(raw);
      const phases = Array.isArray(parsed.phases)
        ? parsed.phases
            .filter((p): p is PlanPhase => Boolean(p?.id && p?.kind && p?.targetSheet))
            .map((p) => ({
              id: String(p.id),
              kind: String(p.kind),
              targetSheet: String(p.targetSheet),
              dependsOn: Array.isArray(p.dependsOn) ? p.dependsOn.map(String) : [],
              ...(Array.isArray(p.repeatFor) && p.repeatFor.length > 0
                ? { repeatFor: p.repeatFor.map(String) }
                : {}),
            }))
        : [];
      const clarificationsNeeded = Array.isArray(parsed.clarificationsNeeded)
        ? parsed.clarificationsNeeded.map(String).filter(Boolean)
        : [];
      const confidence =
        parsed.confidence === 'high' || parsed.confidence === 'medium' || parsed.confidence === 'low'
          ? parsed.confidence
          : 'medium';
      return { phases, clarificationsNeeded, confidence, reasoning: String(parsed.reasoning ?? '') };
    } catch (error: unknown) {
      this.logger.warn(
        `Two-pass planning: coarse JSON parse error: ${error instanceof Error ? error.message : String(error)}. Raw snippet: ${this.clip(raw)}`,
      );
      return null;
    }
  }

  private writePlannerFileLog(args: {
    correlationId: string;
    model: string;
    durationMs: number;
    success: boolean;
    error?: string;
    prompt: string;
    context: WorkbookContext;
    history: { role: string; content: string }[];
    promptContext?: string;
    routerAssumption?: string;
    userMessage: string;
    systemPrompt: string;
    raw: string;
    parsed: PlannerOutput;
    fallback: boolean;
    retried: boolean;
  }): void {
    if (!this.plannerFileLogger) return;

    const userTrunc = truncateForPlannerLog(args.userMessage);
    const rawTrunc = truncateForPlannerLog(args.raw);
    const includeSystem = this.plannerFileLogger.shouldLogFullPrompts();

    this.plannerFileLogger.logPlanner({
      correlationId: args.correlationId,
      model: args.model,
      durationMs: args.durationMs,
      success: args.success,
      ...(args.error ? { error: args.error } : {}),
      input: {
        prompt: args.prompt,
        ...(args.routerAssumption ? { routerAssumption: args.routerAssumption } : {}),
        userMessage: userTrunc.value,
        ...(userTrunc.truncated ? { userMessageTruncated: true } : {}),
        historyLength: args.history.length,
        sheets: args.context.sheets.map((s) => s.name),
        activeSheet: args.context.activeSheetName,
        hasPromptContext: Boolean(args.promptContext?.trim()),
        ...(includeSystem ? { systemPrompt: args.systemPrompt } : {}),
      },
      output: {
        raw: rawTrunc.value,
        ...(rawTrunc.truncated ? { rawTruncated: true } : {}),
        parsed: args.parsed,
        fallback: args.fallback,
        retried: args.retried,
      },
    });
  }

  /**
   * Spec 22 Bug 1: multi-clause "X and Y" write requests must not collapse to one subtask.
   * If the model drops a clause, ask for clarification rather than shipping a half-plan.
   */
  ensureMultiClauseCoverage(prompt: string, plan: PlannerOutput): PlannerOutput {
    if (plan.clarificationsNeeded.length > 0) return plan;
    if (plan.subtasks.length >= 2) {
      return this.ensureDestructiveDependsOnAnnotate(plan);
    }

    const clauses = splitWriteClauses(prompt);
    if (clauses.length < 2) return plan;

    const covered = clauses.filter((clause) =>
      plan.subtasks.some((subtask) => clauseLikelyCovered(clause, subtask.description)),
    );
    if (covered.length >= 2) {
      return this.ensureDestructiveDependsOnAnnotate(plan);
    }

    const missing = clauses.filter(
      (clause) => !plan.subtasks.some((subtask) => clauseLikelyCovered(clause, subtask.description)),
    );
    this.logger?.warn(
      `Planner multi-clause gap: ${missing.length} clause(s) missing from ${plan.subtasks.length} subtask(s)`,
    );
    return {
      ...plan,
      confidence: 'low',
      clarificationsNeeded: [
        ...plan.clarificationsNeeded,
        buildMultiClauseClarification(clauses, missing, plan.subtasks.length),
      ],
      reasoning: `${plan.reasoning} [Spec 22: incomplete multi-clause decomposition]`.trim(),
    };
  }

  /** Prefer annotate/filter before DELETE_COLUMN when both appear in the plan. */
  private ensureDestructiveDependsOnAnnotate(plan: PlannerOutput): PlannerOutput {
    const deleteIdx = plan.subtasks.findIndex((s) =>
      /\bdelete\b.*\bcolumn\b|\bDELETE_COLUMN\b/i.test(s.description) ||
      s.suggestedActionType === 'DELETE_COLUMN',
    );
    const annotateIdx = plan.subtasks.findIndex((s) =>
      /\bremark|priority|unpaid|set\b.*\bwhere\b|SET_MATCHING_ROWS/i.test(s.description) ||
      s.suggestedActionType === 'SET_MATCHING_ROWS',
    );
    if (deleteIdx < 0 || annotateIdx < 0 || deleteIdx === annotateIdx) return plan;

    const annotate = plan.subtasks[annotateIdx]!;
    const del = plan.subtasks[deleteIdx]!;
    if (del.dependsOn.includes(annotate.id)) return plan;

    const subtasks = plan.subtasks.map((s, i) => {
      if (i !== deleteIdx) return s;
      return {
        ...s,
        dependsOn: Array.from(new Set([...s.dependsOn, annotate.id])),
      };
    });
    return { ...plan, subtasks };
  }

  private recordWorkflowNode(
    correlationId: string,
    startedAt: number,
    success: boolean,
    prompt: string,
    context: WorkbookContext,
    parsed: PlannerOutput,
    model: string,
    error?: string,
  ): void {
    this.workflowTrace?.appendNode(correlationId, {
      id: 'planner',
      type: 'planner',
      label: 'Planner',
      status: success ? 'success' : 'failed',
      startedAt: new Date(startedAt),
      endedAt: new Date(),
      durationMs: Date.now() - startedAt,
      input: {
        prompt,
        sheets: context.sheets.map((s) => s.name),
        activeSheet: context.activeSheetName,
      },
      output: {
        subtaskCount: parsed.subtasks.length,
        subtasks: parsed.subtasks.map((s) => ({
          id: s.id,
          description: s.description,
          targetSheet: s.targetSheet,
          dependsOn: s.dependsOn,
        })),
        clarificationsNeeded: parsed.clarificationsNeeded,
        confidence: parsed.confidence,
        reasoning: parsed.reasoning,
        ...(error ? { error } : {}),
      },
      meta: { model, agent: 'planner' },
    });
  }

  private tryParsePlanner(raw: string, correlationId: string, model: string): PlannerOutput | null {
    try {
      const parsed = parseAgentJson<Partial<PlannerOutput>>(raw);
      return this.normalizePlannerOutput(parsed);
    } catch (error: unknown) {
      this.structuredLogger.warnParseFailure(
        correlationId,
        'planner',
        model,
        raw,
        error instanceof Error ? error.message : String(error),
      );
      this.logger.warn(
        `Planner JSON parse error: ${error instanceof Error ? error.message : String(error)}. Raw snippet: ${this.clip(raw)}`,
      );
      return null;
    }
  }

  private normalizePlannerOutput(parsed: Partial<PlannerOutput>): PlannerOutput {
    const rawSubtasks = Array.isArray(parsed.subtasks) ? parsed.subtasks : [];
    // Dropping a malformed subtask silently is how a truncated plan passed as
    // complete: the model's last subtask was cut mid-string (no targetSheet),
    // this filter removed it, and the caller saw a clean, shorter plan with no
    // indication anything was missing. Count and report what we discard.
    const droppedSubtasks = rawSubtasks.filter(
      (s) => !(s?.id && s?.description && s?.targetSheet),
    );
    if (droppedSubtasks.length > 0) {
      this.logger.warn(
        `Planner emitted ${rawSubtasks.length} subtask(s) but ${droppedSubtasks.length} were ` +
          `malformed and dropped (missing id/description/targetSheet) — ` +
          `ids: [${droppedSubtasks.map((s) => String(s?.id ?? '(no id)')).join(', ')}]. ` +
          `This usually means the plan was cut off mid-generation; the remaining plan may be incomplete.`,
      );
    }
    const subtasks = rawSubtasks
          .filter((s): s is SubTask => Boolean(s?.id && s?.description && s?.targetSheet))
          .map((s) => {
            const subtask: SubTask = {
              id: String(s.id),
              description: String(s.description),
              targetSheet: String(s.targetSheet),
              dependsOn: Array.isArray(s.dependsOn) ? s.dependsOn.map(String) : [],
              estimatedActions:
                typeof s.estimatedActions === 'number' ? s.estimatedActions : 1,
            };
            if (typeof s.suggestedActionType === 'string' && s.suggestedActionType.trim()) {
              subtask.suggestedActionType = s.suggestedActionType.trim();
            }
            return subtask;
          });

    const clarificationsNeeded = Array.isArray(parsed.clarificationsNeeded)
      ? parsed.clarificationsNeeded.map(String).filter(Boolean)
      : [];

    const confidence =
      parsed.confidence === 'high' ||
      parsed.confidence === 'medium' ||
      parsed.confidence === 'low'
        ? parsed.confidence
        : 'medium';

    return {
      subtasks,
      clarificationsNeeded,
      confidence,
      reasoning: String(parsed.reasoning ?? ''),
    };
  }

  /**
   * Drops subtasks whose `dependsOn` can never be satisfied — directly or
   * transitively — because they reference a subtask id that does not exist in
   * this plan at all, and reports what was dropped via `clarificationsNeeded`
   * (the existing "proceeded under an assumption" channel, so this needs no
   * new plumbing through Orchestrator/ConversationService).
   *
   * Live incident (Sept 8, 2026): a 19-subtask plan's raw response ended
   * mid-generation while listing "July" (subtask s14) — the model's own
   * closing brackets made the JSON syntactically VALID, so `outcome.truncated`
   * was never set and the existing truncation-retry ladder never fired. The
   * result: 14 real subtasks, but s2's `dependsOn` still named `s15`..`s19`
   * (August through December, never generated at all). Nothing caught this —
   * `normalizePlannerOutput` only checks that EACH subtask individually has an
   * id/description/targetSheet, never that `dependsOn` edges resolve. The plan
   * looked complete and confident, `computeExecutionWaves`'s "stranded" wave
   * fallback dumped the unsatisfiable subtasks (s2-s7, the Main-sheet totals,
   * KPI band, and chart) into a final wave anyway, and their formulas —
   * `=SUM(August!G:G)` etc. — landed in the user's real workbook as 18 `#REF!`
   * errors on sheets that were never created.
   *
   * Pruning (not retrying) on discovery is the safer failure mode: a request
   * this large already risks exhausting even the last-resort token ceiling
   * (TASKS.md #170), so throwing the WHOLE plan away to retry would trade a
   * partial, honestly-reported build for a real chance of no build at all.
   * Delivering the 12 sheets that ARE fully specified, while openly saying 5
   * months and the totals that depended on them could not be planned this
   * time, is the CODEBASE_ANALYSIS.md §3.7 rule: partial-and-honest beats
   * complete-looking-and-wrong.
   */
  private pruneUnsatisfiableSubtasks(output: PlannerOutput): PlannerOutput {
    const validIds = new Set(output.subtasks.map((s) => s.id));
    const unsatisfiable = new Set<string>();

    // Fixed point: a subtask depending on an already-unsatisfiable subtask is
    // itself unsatisfiable, even if its OWN dependsOn ids all otherwise exist —
    // s3 depends only on s2, but s2 depends on the missing s15, so s3 can never
    // safely run either.
    let changed = true;
    while (changed) {
      changed = false;
      for (const subtask of output.subtasks) {
        if (unsatisfiable.has(subtask.id)) continue;
        const hasUnsatisfiableDep = subtask.dependsOn.some(
          (dep) => !validIds.has(dep) || unsatisfiable.has(dep),
        );
        if (hasUnsatisfiableDep) {
          unsatisfiable.add(subtask.id);
          changed = true;
        }
      }
    }

    if (unsatisfiable.size === 0) return output;

    const dropped = output.subtasks.filter((s) => unsatisfiable.has(s.id));
    const kept = output.subtasks.filter((s) => !unsatisfiable.has(s.id));

    this.logger.error(
      `Planner plan referenced ${unsatisfiable.size} subtask id(s) that were never generated ` +
        `(dangling dependsOn — likely a truncation the provider did not flag) — pruning ` +
        `unsatisfiable subtask(s) [${dropped.map((s) => s.id).join(', ')}] rather than shipping ` +
        `formulas/actions that would reference sheets or ranges nothing ever created. ` +
        `${kept.length} of ${output.subtasks.length} subtasks remain deliverable.`,
    );

    const droppedSheets = [...new Set(dropped.map((s) => s.targetSheet))];
    const note =
      `I could not fully plan ${dropped.length} step${dropped.length === 1 ? '' : 's'} of this ` +
      `request (affecting ${droppedSheets.join(', ')}) — the response was cut off before they were ` +
      `fully specified. I've applied the rest; ask again to add the missing part once you see what's here.`;

    return {
      ...output,
      subtasks: kept,
      clarificationsNeeded: [...output.clarificationsNeeded, note],
    };
  }

  /** Exposed for unit tests — preserves suggestedActionType and required fields. */
  normalizePlannerOutputForTest(parsed: Partial<PlannerOutput>): PlannerOutput {
    return this.normalizePlannerOutput(parsed);
  }

  /**
   * TASKS.md #230 — a subtask that references `Lists!$B$3:$B$20` (a dropdown
   * source) or `January!G:G` (a formula) needs that sheet to exist. When no
   * subtask creates it and the workbook lacks it, add a create subtask the
   * referencing subtasks depend on. Live: January's dropdowns pointed at a
   * Lists sheet nothing ever created.
   */
  private planReferencedSheets(output: PlannerOutput, context: WorkbookContext): PlannerOutput {
    const { plan, added } = ensureReferencedSheetsPlanned(output, context);
    if (added.length > 0) {
      this.logger.warn(
        `Planner plan referenced sheet(s) no subtask creates — added create subtask(s) for: ${added.join(', ')}.`,
      );
    }
    return plan;
  }

  /** Exposed for unit tests. */
  pruneUnsatisfiableSubtasksForTest(output: PlannerOutput): PlannerOutput {
    return this.pruneUnsatisfiableSubtasks(output);
  }

  /**
   * Only return fallbacks that are useful structured plans — never a single
   * subtask whose description is the raw user message.
   */
  private tryUsefulFallbackPlan(prompt: string, context: WorkbookContext): PlannerOutput | null {
    const activeSheet = context.activeSheetName || 'Sheet1';
    const sheet = context.sheets.find((s) => s.name === activeSheet);
    const hasValues = sheet?.values.some((row) =>
      row?.some((cell) => cell !== null && cell !== '' && String(cell).trim() !== ''),
    );
    const hasHeaders = (sheet?.values[0] ?? []).some(
      (cell) => cell !== null && cell !== '' && String(cell).trim() !== '',
    );
    const isEmpty = !sheet || (!hasValues && !hasHeaders && sheet.rowCount === 0);

    if (isEmpty) {
      return {
        subtasks: [],
        clarificationsNeeded: [
          'I could not read sheet data (the workbook may be empty or still loading). Please try again after the sheet has data, or specify which column to sort by.',
        ],
        confidence: 'low',
        reasoning: 'Fallback — empty or unreadable workbook context',
      };
    }

    const compoundSubtasks = buildCompoundFallbackSubtasks(prompt, context);
    if (compoundSubtasks) {
      return {
        subtasks: compoundSubtasks,
        clarificationsNeeded: [],
        confidence: 'low',
        reasoning: 'Fallback compound plan — create sheet then sort',
      };
    }

    return null;
  }

  private clip(value: string, max = 400): string {
    const normalized = value.replace(/\s+/g, ' ').trim();
    return normalized.length <= max ? normalized : `${normalized.slice(0, max)}...`;
  }
}

/**
 * Verbs that make a fragment a COMMAND rather than a continuation of a question.
 * Kept in sync with the write-intent guard's verb list by intent, not by import —
 * this one is about "is this fragment an instruction", not "does this prompt write".
 */
const CLAUSE_COMMAND_VERB =
  /\b(sort|filter|delete|remove|insert|add|copy|move|bold|highlight|colou?r|format|merge|split|fill|clear|rename|hide|unhide|freeze|protect|create|build|generate|apply|replace|update|change|set|mark|label|flag)\b/i;

/**
 * Fragments that are questions, not commands — a clause starting this way is the
 * tail of an interrogative sentence even when it contains a command-shaped verb.
 * "…and what they add up to" is one clause with the question before it, not two.
 */
const CLAUSE_IS_QUESTION =
  /^(what|which|how|why|when|where|who|whose|whether|if|do|does|did|is|are|was|were|can|could|should|would)\b/i;

/**
 * Split compound write prompts on and/then/also connectors.
 *
 * Task #91 (2026-08-27): splitting on a bare connector treated "…how many invoices
 * are pending payment **and** what they add up to" as two clauses — a conjunction
 * joining two objects of one question, not two instructions — and the coverage check
 * then raised a false multi-clause gap on a read-only prompt. A fragment now counts
 * as a clause only if it reads as a command: it must contain a command verb and must
 * not open like a question.
 */
export function splitWriteClauses(prompt: string): string[] {
  const normalized = prompt.replace(/\s+/g, ' ').trim();
  if (!normalized) return [];
  const parts = normalized
    .split(/\s+(?:and|, and|then|also)\s+/i)
    .map((p) => p.trim())
    .filter((p) => p.length >= 8);
  if (parts.length < 2) return [];

  const commandClauses = parts.filter(
    (p) => CLAUSE_COMMAND_VERB.test(p) && !CLAUSE_IS_QUESTION.test(p),
  );
  // Only a genuine multi-COMMAND prompt can have an incomplete decomposition.
  return commandClauses.length >= 2 ? commandClauses : [];
}

/**
 * Clarification text built from the user's ACTUAL clauses.
 *
 * Task #91 (2026-08-27): this string used to be hardcoded to the Spec 22 scenario
 * it was written for — "Should I handle both — annotate/filter first, then any
 * column deletion?" — and was emitted verbatim regardless of prompt. A user who
 * never mentioned deleting anything was asked about column deletion, which reads as
 * the agent hallucinating destructive intent. Never mention an operation the user
 * did not ask for.
 */
export function buildMultiClauseClarification(
  clauses: string[],
  missing: string[],
  plannedCount: number,
): string {
  const quote = (c: string) => `"${c.length > 60 ? `${c.slice(0, 59)}…` : c}"`;
  const missingList = (missing.length > 0 ? missing : clauses).map(quote).join(' and ');
  const stepWord = plannedCount === 1 ? 'step' : 'steps';
  return (
    `Your request looks like ${clauses.length} separate instructions, but I only ` +
    `planned ${plannedCount} ${stepWord}. I don't have a plan for ${missingList}. ` +
    `Should I handle everything you asked for?`
  );
}

export function clauseLikelyCovered(clause: string, description: string): boolean {
  const clauseTokens = significantTokens(clause);
  const descTokens = new Set(significantTokens(description));
  if (clauseTokens.length === 0) return false;
  const overlap = clauseTokens.filter((t) => descTokens.has(t)).length;
  return overlap >= Math.min(2, clauseTokens.length);
}

function significantTokens(text: string): string[] {
  const stop = new Set([
    'the', 'a', 'an', 'to', 'of', 'in', 'on', 'for', 'with', 'and', 'or', 'by',
    'from', 'that', 'this', 'into', 'add', 'set', 'make', 'please', 'column',
  ]);
  return text
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter((t) => t.length > 2 && !stop.has(t));
}
