import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import { EXECUTOR_SYSTEM_PROMPT, buildExecutorUserMessage } from './prompts/executor.prompt';
import { normalizeExecutorOutput } from './utils/normalize-executor-output.util';
import { parseExecutorPayload } from './utils/parse-agent-json.util';
import {
  buildDeterministicSubtaskActions,
  maybeMarkSubtaskComplete,
} from './utils/compound-action.util';
import { buildSortFallbackAction } from './utils/sort-action.util';
import { Action, ExecutorOutput, SubTask, WorkbookContext } from './types/agent.types';
import { StepRetryContext } from './types/verifier.types';
import { StepRetryExhaustedError } from './errors';
import { StructuredLogger } from './logging/structured-logger';
import { isExecutorBlockedSignal } from './utils/verifier-partial-parse.util';

const JSON_RETRY_SUFFIX =
  '\n\nIMPORTANT: Your previous response was not valid JSON. Reply with ONLY a single JSON object matching the schema — no markdown fences, no commentary.';

@Injectable()
export class ExecutorAgent {
  private readonly logger = new Logger(ExecutorAgent.name);

  constructor(
    private readonly llm: OpenRouterService,
    private readonly config: AppConfigService,
    private readonly structuredLogger: StructuredLogger = new StructuredLogger(),
  ) {}

  get modelName(): string {
    return this.config.openRouterModelHigh;
  }

  /** Single LLM invocation per call — agentic loop owns multi-step iteration and virtual apply. */
  async execute(
    subtask: SubTask,
    context: WorkbookContext,
    previousActions: Action[] = [],
    correlationId = `req_${Date.now()}`,
  ): Promise<ExecutorOutput> {
    const startedAt = Date.now();
    const model = this.modelName;
    const userMessage = buildExecutorUserMessage(subtask, context, previousActions);

    let raw = await this.llm.complete({
      systemPrompt: EXECUTOR_SYSTEM_PROMPT,
      userMessage,
      model,
      temperature: 0.1,
      maxTokens: 2000,
    });
    this.structuredLogger.debugRawResponse(correlationId, 'executor', model, raw);

    let result = this.tryParseExecutor(raw, subtask);
    let parsedOnFirstAttempt = true;
    if (!result) {
      parsedOnFirstAttempt = false;
      this.structuredLogger.warnParseFailure(
        correlationId,
        'executor',
        model,
        raw,
        'First parse attempt failed',
      );
      this.logger.warn(`Executor JSON parse failed — retrying once. Raw snippet: ${this.clip(raw)}`);
      raw = await this.llm.complete({
        systemPrompt: EXECUTOR_SYSTEM_PROMPT,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        model,
        temperature: 0.05,
        maxTokens: 2000,
      });
      this.structuredLogger.debugRawResponse(correlationId, 'executor', model, raw);
      result = this.tryParseExecutor(raw, subtask);
    }

    if (result) {
      if (result.actions.length === 0 && !result.isDone) {
        // Spec 17 Bug C: never paper over an honest block with SORT_RANGE.
        if (isExecutorBlockedSignal(result.nextStep)) {
          this.logger.warn(
            `Executor blocked (no sort fallback): ${this.clip(result.nextStep ?? '', 300)}`,
          );
          this.structuredLogger.logAgentEvent({
            correlationId,
            agent: 'executor',
            model,
            durationMs: Date.now() - startedAt,
            success: false,
            tokenUsage: this.structuredLogger.estimateTokens(raw),
            rawResponse: raw,
            parsedResponse: result,
            error: 'Executor blocked',
          });
          return { ...result, parsedOnFirstAttempt };
        }
        return { ...this.applySortFallback(subtask, context, result), parsedOnFirstAttempt };
      }
      result = maybeMarkSubtaskComplete(result, subtask);
      this.logger.log(
        `Executor produced ${result.actions.length} actions for subtask ${subtask.id}`,
      );
      this.structuredLogger.logAgentEvent({
        correlationId,
        agent: 'executor',
        model,
        durationMs: Date.now() - startedAt,
        success: true,
        tokenUsage: this.structuredLogger.estimateTokens(raw),
        rawResponse: raw,
        parsedResponse: result,
      });
      return { ...result, parsedOnFirstAttempt };
    }

    this.logger.warn(
      `Executor returned invalid JSON after retry — using fallback. Snippet: ${this.clip(raw)}`,
    );
    this.structuredLogger.warnParseFailure(
      correlationId,
      'executor',
      model,
      raw,
      'Parse failed after retry',
    );
    const fallback = this.buildFailureFallback(subtask, context);
    this.structuredLogger.logAgentEvent({
      correlationId,
      agent: 'executor',
      model,
      durationMs: Date.now() - startedAt,
      success: false,
      tokenUsage: this.structuredLogger.estimateTokens(raw),
      rawResponse: raw,
      parsedResponse: fallback,
      error: 'Executor JSON parse failed after retry',
    });
    return { ...fallback, parsedOnFirstAttempt: false };
  }

  async retryStep(
    retryContext: StepRetryContext,
    context: WorkbookContext,
    previousActions: Action[] = [],
    correlationId = `req_${Date.now()}`,
  ): Promise<ExecutorOutput> {
    const { originalStep, attempt, maxAttempts, verifierFeedback } = retryContext;

    if (attempt > maxAttempts) {
      throw new StepRetryExhaustedError(
        `Step "${originalStep.id}" failed after ${maxAttempts} retry attempts.`,
        { step: originalStep, attempts: maxAttempts },
      );
    }

    this.logger.warn(
      `Executor retry step ${originalStep.id} attempt ${attempt}/${maxAttempts}: ${this.clip(verifierFeedback, 300)}`,
    );

    const retryAwareContext: WorkbookContext = {
      ...context,
      verifierFeedback,
    };

    return this.execute(originalStep, retryAwareContext, previousActions, correlationId);
  }

  private tryParseExecutor(raw: string, subtask: SubTask): ExecutorOutput | null {
    const parsed = parseExecutorPayload(raw);
    if (!parsed) return null;
    return normalizeExecutorOutput(parsed, subtask);
  }

  private applySortFallback(
    subtask: SubTask,
    context: WorkbookContext,
    partial: ExecutorOutput,
  ): ExecutorOutput {
    const sortAction = buildSortFallbackAction(subtask, context);
    if (!sortAction) return partial;

    this.logger.log(`Executor sort fallback for column "${sortAction.columnName}"`);
    return {
      subtaskId: subtask.id,
      actions: [sortAction],
      isDone: true,
      nextStep: partial.nextStep,
    };
  }

  private buildFailureFallback(subtask: SubTask, context: WorkbookContext): ExecutorOutput {
    const deterministic = buildDeterministicSubtaskActions(subtask, context);
    if (deterministic) {
      this.logger.log(
        `Executor JSON fallback — built ${deterministic.actions.length} deterministic action(s)`,
      );
      return deterministic;
    }

    const sortAction = buildSortFallbackAction(subtask, context);
    if (sortAction) {
      this.logger.log(`Executor JSON fallback — built SORT_RANGE for "${sortAction.columnName}"`);
      return {
        subtaskId: subtask.id,
        actions: [sortAction],
        isDone: true,
      };
    }

    return {
      subtaskId: subtask.id,
      actions: [],
      isDone: false,
      nextStep:
        'I could not produce valid actions for this step. Try specifying the column name or a simpler instruction.',
    };
  }

  private clip(value: string, max = 400): string {
    const normalized = value.replace(/\s+/g, ' ').trim();
    return normalized.length <= max ? normalized : `${normalized.slice(0, max)}...`;
  }
}
