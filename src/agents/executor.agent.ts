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

const JSON_RETRY_SUFFIX =
  '\n\nIMPORTANT: Your previous response was not valid JSON. Reply with ONLY a single JSON object matching the schema — no markdown fences, no commentary.';

@Injectable()
export class ExecutorAgent {
  private readonly logger = new Logger(ExecutorAgent.name);

  constructor(
    private readonly llm: OpenRouterService,
    private readonly config: AppConfigService,
  ) {}

  /** Single LLM invocation per call — agentic loop owns multi-step iteration and virtual apply. */
  async execute(
    subtask: SubTask,
    context: WorkbookContext,
    previousActions: Action[] = [],
  ): Promise<ExecutorOutput> {
    const userMessage = buildExecutorUserMessage(subtask, context, previousActions);

    let raw = await this.llm.complete({
      systemPrompt: EXECUTOR_SYSTEM_PROMPT,
      userMessage,
      model: this.config.openRouterModelHigh,
      temperature: 0.1,
      maxTokens: 2000,
    });

    let result = this.tryParseExecutor(raw, subtask);
    if (!result) {
      this.logger.warn('Executor JSON parse failed — retrying once');
      raw = await this.llm.complete({
        systemPrompt: EXECUTOR_SYSTEM_PROMPT,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        model: this.config.openRouterModelHigh,
        temperature: 0.05,
        maxTokens: 2000,
      });
      result = this.tryParseExecutor(raw, subtask);
    }

    if (result) {
      if (result.actions.length === 0 && !result.isDone) {
        return this.applySortFallback(subtask, context, result);
      }
      result = maybeMarkSubtaskComplete(result, subtask);
      this.logger.log(
        `Executor produced ${result.actions.length} actions for subtask ${subtask.id}`,
      );
      return result;
    }

    this.logger.warn(
      `Executor returned invalid JSON after retry — using fallback. Snippet: ${this.clip(raw)}`,
    );
    return this.buildFailureFallback(subtask, context);
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
