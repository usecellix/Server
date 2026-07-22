import { Injectable, Logger, Optional } from '@nestjs/common';
import { PlannerFileLoggerService } from '../common/logging/planner-file-logger.service';
import { truncateForPlannerLog } from '../common/logging/planner-file-logger.util';
import { AppConfigService } from '../config/app-config.service';
import { PLANNER_RULES_ADDITION } from '../excel-ai/prompt/cellix-system-prompt';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import {
  PLANNER_EXHAUSTED_USER_MESSAGE,
  PlannerExhaustedError,
} from './errors';
import { PLANNER_SYSTEM_PROMPT, buildPlannerUserMessage } from './prompts/planner.prompt';
import { parseAgentJson } from './utils/parse-agent-json.util';
import { buildCompoundFallbackSubtasks } from './utils/compound-action.util';
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

    let raw = await this.llm.complete({
      ...completeOpts,
      userMessage,
      temperature: 0.2,
    });
    this.structuredLogger.debugRawResponse(correlationId, 'planner', model, raw);

    let retried = false;
    let lastResort = false;
    let parsed = this.tryParsePlanner(raw, correlationId, model);
    if (!parsed) {
      retried = true;
      this.logger.warn(`Planner JSON parse failed — retrying once. Raw snippet: ${this.clip(raw)}`);
      raw = await this.llm.complete({
        ...completeOpts,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        temperature: 0.1,
      });
      this.structuredLogger.debugRawResponse(correlationId, 'planner', model, raw);
      parsed = this.tryParsePlanner(raw, correlationId, model);
    }

    if (!parsed) {
      lastResort = true;
      this.logger.warn(
        `Planner still empty/unparseable — last-resort retry with maxTokens=${PLANNER_LAST_RESORT_MAX_TOKENS}`,
      );
      raw = await this.llm.complete({
        ...completeOpts,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        temperature: 0.1,
        maxTokens: PLANNER_LAST_RESORT_MAX_TOKENS,
        reasoningMaxTokens: Math.min(PLANNER_REASONING_MAX_TOKENS, 768),
      });
      this.structuredLogger.debugRawResponse(correlationId, 'planner', model, raw);
      parsed = this.tryParsePlanner(raw, correlationId, model);
    }

    if (parsed) {
      this.logger.log(
        `Planner produced ${parsed.subtasks.length} subtasks, confidence: ${parsed.confidence}`,
      );
      this.structuredLogger.logAgentEvent({
        correlationId,
        agent: 'planner',
        model,
        durationMs: Date.now() - startedAt,
        success: true,
        tokenUsage: this.structuredLogger.estimateTokens(raw),
        rawResponse: raw,
        parsedResponse: parsed,
      });
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
        parsed,
        fallback: false,
        retried: retried || lastResort,
      });
      return parsed;
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
    const subtasks = Array.isArray(parsed.subtasks)
      ? parsed.subtasks
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
          })
      : [];

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
