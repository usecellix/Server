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
  PLANNER_EXHAUSTED_USER_MESSAGE,
  PlannerExhaustedError,
} from './errors';
import { PLANNER_SYSTEM_PROMPT, buildPlannerUserMessage } from './prompts/planner.prompt';
import { parseAgentJson } from './utils/parse-agent-json.util';
import { buildCompoundFallbackSubtasks } from './utils/compound-action.util';
import { ensureNumberFormatPlanSafety } from './utils/preserve-number-format.util';
import {
  PLANNER_LAST_RESORT_MAX_TOKENS,
  PLANNER_REASONING_MAX_TOKENS,
  resolvePlannerMaxTokens,
} from './utils/planner-token-budget.util';
import { PlannerOutput, SubTask, WorkbookContext } from './types/agent.types';
import { StructuredLogger } from './logging/structured-logger';

const JSON_RETRY_SUFFIX =
  '\n\nIMPORTANT: Your previous response was not valid JSON. Reply with ONLY a single JSON object matching the schema — no markdown fences, no commentary.';

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
  ): Promise<PlannerOutput> {
    const startedAt = Date.now();
    const model = this.config.openRouterModelHigh;
    const systemPrompt = PLANNER_SYSTEM_PROMPT + PLANNER_RULES_ADDITION;
    let userMessage = buildPlannerUserMessage(prompt, context, history, promptContext);
    if (routerAssumption) {
      userMessage = `[Router assumption: ${routerAssumption}]\n\n${userMessage}`;
    }

    const maxTokens = resolvePlannerMaxTokens(complexity);
    const completeOpts = {
      systemPrompt,
      model,
      maxTokens,
      reasoningEffort: 'low' as const,
      reasoningMaxTokens: PLANNER_REASONING_MAX_TOKENS,
    };

    // A plan cut off by the token budget frequently still parses — the model
    // closes the JSON it has emitted so far — so parse success alone is NOT
    // evidence the plan is complete. Treat truncation as a failure to be
    // retried on a bigger budget, exactly like a parse failure.
    const outcome: LlmCompletionOutcome = {};
    let raw = await this.llm.complete({
      ...completeOpts,
      userMessage,
      temperature: 0.2,
      outcome,
    });
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
      const covered = ensureNumberFormatPlanSafety(
        prompt,
        this.ensureMultiClauseCoverage(prompt, parsed),
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

  /** Exposed for unit tests — preserves suggestedActionType and required fields. */
  normalizePlannerOutputForTest(parsed: Partial<PlannerOutput>): PlannerOutput {
    return this.normalizePlannerOutput(parsed);
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
