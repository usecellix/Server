import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../../config/app-config.service';
import { LLMTier } from '../../types/cellix.types';
import { LlmRequestError } from '../errors/llm-request.error';
import { ModelRouter } from '../llm/model-router';
import { extractChatContent } from '../utils/extract-chat-content.util';
import { isReasoningMandatoryError } from '../utils/reasoning-mandatory.util';
import {
  computeReasoningMaxTokens,
  ensureBudgetForReasoning,
  escalatedRetryBudget,
} from '../utils/reasoning-budget.util';

export type OpenRouterChatMessage = {
  role: 'system' | 'user' | 'assistant';
  content: string;
};

export type LlmUsage = {
  promptTokens?: number;
  completionTokens?: number;
  totalTokens?: number;
  reasoningTokens?: number;
};

/**
 * Mutable out-param for `complete()`. `truncated` is true when the provider
 * stopped because the completion hit its token budget rather than because the
 * model finished — the case that silently produces short-but-parseable JSON.
 */
export type LlmCompletionOutcome = {
  finishReason?: string | null;
  truncated?: boolean;
};

/** finishReason values that mean "cut off by the token budget", across providers. */
const TRUNCATION_FINISH_REASONS = new Set(['length', 'max_tokens', 'max_output_tokens']);

export function isTruncationFinishReason(reason: string | null | undefined): boolean {
  return typeof reason === 'string' && TRUNCATION_FINISH_REASONS.has(reason.toLowerCase());
}

export type LlmCallTelemetry = {
  provider?: string;
  model?: string;
  modelTier?: string;
  usage?: LlmUsage;
  complexityScore?: number;
  routingRationale?: string;
  fallbackUsed?: boolean;
  estimatedCostUsd?: number;
};

type ReasoningEffort = 'none' | 'minimal' | 'low' | 'medium' | 'high';

type ChatMessage = {
  role: 'system' | 'user' | 'assistant';
  content: string;
};

type ChatCompletionResult = {
  choices?: Array<{
    finishReason?: string | null;
    message?: {
      content?: unknown;
      refusal?: string | null;
    };
  }>;
  usage?: {
    completionTokens?: number;
    promptTokens?: number;
    totalTokens?: number;
    completionTokensDetails?: {
      reasoningTokens?: number | null;
    };
  };
};

type OpenRouterClient = InstanceType<Awaited<typeof import('@openrouter/sdk')>['OpenRouter']>;

@Injectable()
export class OpenRouterService {
  private readonly logger = new Logger(OpenRouterService.name);

  constructor(
    private readonly config: AppConfigService,
    private readonly modelRouter: ModelRouter,
  ) {}

  isConfigured(): boolean {
    return Boolean(this.config.openRouterApiKey);
  }

  async complete(opts: {
    systemPrompt: string;
    userMessage: string;
    model?: string;
    tier?: LLMTier;
    temperature?: number;
    maxTokens?: number;
    reasoningEffort?: ReasoningEffort;
    /** Cap reasoning tokens independently when the provider supports it. */
    reasoningMaxTokens?: number;
    responseFormat?: 'json_object' | 'text';
    /**
     * Optional out-param: populated with the provider's finishReason for the
     * completion actually returned. A caller that parses structured output MUST
     * check this — a model that hits its token cap mid-JSON often emits the
     * closing brackets anyway, so the result parses cleanly and is silently
     * short. (Real incident: a Planner run truncated mid-subtask, still parsed,
     * and the missing subtasks were never noticed by any downstream check.)
     */
    outcome?: LlmCompletionOutcome;
  }): Promise<string> {
    const apiKey = this.config.openRouterApiKey;
    if (!apiKey) {
      throw new LlmRequestError(503, 'OpenRouter not configured');
    }

    const model =
      opts.model ??
      (opts.tier === 'low'
        ? this.config.openRouterModelLow
        : opts.tier === 'high'
          ? this.config.openRouterModelHigh
          : opts.tier === 'medium'
            ? this.config.openRouterModelMedium
            : this.config.openRouterModelMedium);
    const budget = opts.maxTokens ?? 1500;
    // Reasoning models share the completion budget between thinking and output.
    // Raise starved budgets and always cap reasoning so content has room. See F5.
    const completionBudget = ensureBudgetForReasoning(model, budget);
    const reasoningCap =
      opts.reasoningMaxTokens ?? computeReasoningMaxTokens(model, completionBudget);
    const requestedEffort = opts.reasoningEffort ?? 'low';

    try {
      const { OpenRouter } = await import('@openrouter/sdk');
      const client = new OpenRouter({
        apiKey,
        httpReferer: this.config.openRouterHttpReferer,
        appTitle: 'Cellix',
      });

      const baseMessages: ChatMessage[] = [
        { role: 'system', content: opts.systemPrompt },
        { role: 'user', content: opts.userMessage },
      ];

      let response = await this.requestChatCompletion(client, {
        model,
        messages: baseMessages,
        temperature: opts.temperature ?? 0.2,
        maxCompletionTokens: completionBudget,
        reasoningEffort: requestedEffort,
        reasoningMaxTokens: reasoningCap,
        responseFormat: opts.responseFormat ?? 'json_object',
      });

      let text = this.extractCompletionText(response);
      if (!text.trim()) {
        this.logEmptyCompletion(model, response, 'first attempt');
        // The retry MUST differ from the first attempt or it just burns latency
        // twice for the same result (F5: two identical 4096-token runs = 85s).
        // Escalate the budget and tighten the reasoning cap so content has room.
        // Prefer low over none — gpt-5 family rejects effort=none.
        const retryBudget = escalatedRetryBudget(completionBudget);
        const retryReasoningCap =
          computeReasoningMaxTokens(model, retryBudget) ?? reasoningCap;
        this.logger.warn(
          `OpenRouter retrying after empty content (model=${model}, ` +
            `budget ${completionBudget}->${retryBudget}, reasoningCap ${reasoningCap ?? 'none'}->${retryReasoningCap ?? 'none'})`,
        );
        response = await this.requestChatCompletion(client, {
          model,
          messages: baseMessages,
          temperature: opts.temperature ?? 0.2,
          maxCompletionTokens: retryBudget,
          reasoningEffort: 'low',
          reasoningMaxTokens: retryReasoningCap,
          responseFormat: opts.responseFormat ?? 'json_object',
        });
        text = this.extractCompletionText(response);
        if (!text.trim()) {
          this.logEmptyCompletion(model, response, 'retry with reasoning.effort=low');
        }
      }

      if (opts.outcome) {
        const finishReason = response.choices?.[0]?.finishReason ?? null;
        opts.outcome.finishReason = finishReason;
        opts.outcome.truncated = isTruncationFinishReason(finishReason);
        if (opts.outcome.truncated) {
          this.logger.warn(
            `OpenRouter completion truncated by token budget (model=${model}, ` +
              `finishReason=${finishReason}, maxTokens=${completionBudget}) — ` +
              `structured output from this call may parse but be incomplete.`,
          );
        }
      }

      return text;
    } catch (error: unknown) {
      const status = this.extractStatus(error);
      const detail = error instanceof Error ? error.message : 'OpenRouter complete failed';
      this.logger.warn(`OpenRouter complete failed (${status}): ${detail}`);
      throw new LlmRequestError(status, detail);
    }
  }

  async quickCall(systemPrompt: string, userMessage: string): Promise<string> {
    const apiKey = this.config.openRouterApiKey;
    if (!apiKey) {
      return '';
    }

    try {
      const { OpenRouter } = await import('@openrouter/sdk');
      const client = new OpenRouter({
        apiKey,
        httpReferer: this.config.openRouterHttpReferer,
        appTitle: 'Cellix',
      });

      const response = await this.requestChatCompletion(client, {
        model: this.config.openRouterModelLow,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userMessage },
        ],
        temperature: 0,
        maxCompletionTokens: 512,
        reasoningEffort: 'none',
        responseFormat: 'json_object',
      });

      return this.extractCompletionText(response);
    } catch (error: unknown) {
      const detail = error instanceof Error ? error.message : 'OpenRouter quickCall failed';
      this.logger.warn(`OpenRouter quickCall failed: ${detail}`);
      return '';
    }
  }

  /**
   * Stream a chat completion, retrying once on a transient fault.
   *
   * Task #92: the write path previously had no retry at all — one transient error
   * dropped the user into the local rule engine, which refuses structural writes,
   * so a blip meant total failure of the request. The retry is only safe before any
   * token has been yielded; after that a second stream would duplicate output, so
   * `streamChatOnce` reports whether it emitted anything.
   */
  async *streamChat(
    messages: OpenRouterChatMessage[],
    telemetry?: LlmCallTelemetry,
    model = this.config.openRouterModelMedium,
    maxTokens = 4096,
  ): AsyncGenerator<string> {
    const progress = { yielded: false };
    try {
      yield* this.streamChatOnce(messages, telemetry, model, maxTokens, progress);
      return;
    } catch (error: unknown) {
      const status = this.extractStatus(error);
      const detail = error instanceof Error ? error.message : 'OpenRouter request failed';
      // A 402 is deliberately NOT retried (see LlmRequestError.isRetryable) — an
      // empty wallet is still empty a second later, and retrying only doubles the
      // wait before the same failure.
      const retryable =
        new LlmRequestError(status, detail).isRetryable ||
        this.isTransientNetworkError(error);
      if (!retryable || progress.yielded) {
        throw error;
      }
      this.logger.warn(`OpenRouter chat retrying once (${status}): ${detail}`);
    }
    yield* this.streamChatOnce(messages, telemetry, model, maxTokens);
  }

  private async *streamChatOnce(
    messages: OpenRouterChatMessage[],
    telemetry?: LlmCallTelemetry,
    model = this.config.openRouterModelMedium,
    maxTokens = 4096,
    progress?: { yielded: boolean },
  ): AsyncGenerator<string> {
    const apiKey = this.config.openRouterApiKey;
    if (!apiKey) {
      return;
    }

    if (telemetry) {
      telemetry.provider = 'openrouter';
      telemetry.model = model;
    }

    try {
      const { OpenRouter } = await import('@openrouter/sdk');
      const client = new OpenRouter({
        apiKey,
        httpReferer: this.config.openRouterHttpReferer,
        appTitle: 'Cellix',
      });

      // Reasoning models share maxCompletionTokens between thinking and output.
      // Without an explicit cap the model can spend the whole budget reasoning and
      // stream zero content chunks — the F5 incident (85s, empty answer, reported
      // to the user as a JSON parse failure). Cap reasoning to reserve content room.
      const completionBudget = ensureBudgetForReasoning(model, maxTokens);
      const reasoningCap = computeReasoningMaxTokens(model, completionBudget);
      this.logger.log(
        `F5-GUARD streamChat model=${model} maxCompletionTokens=${completionBudget} ` +
          `reasoningCap=${reasoningCap ?? 'none'}`,
      );

      const stream = await client.chat.send({
        chatRequest: {
          model,
          messages,
          stream: true,
          streamOptions: { includeUsage: true },
          temperature: 0.25,
          maxCompletionTokens: completionBudget,
          ...(reasoningCap
            ? // SDK typings may omit max_tokens on reasoning — cast for providers that support it.
              {
                reasoning: { effort: 'low' as ReasoningEffort, max_tokens: reasoningCap } as {
                  effort: ReasoningEffort;
                },
              }
            : {}),
        },
      });

      let streamedChars = 0;
      let lastUsage: LlmUsage | undefined;
      for await (const chunk of stream) {
        const usage = this.normalizeUsage((chunk as { usage?: Record<string, unknown> }).usage);
        if (usage) {
          lastUsage = usage;
          if (telemetry) {
            telemetry.usage = usage;
          }
        }

        const content = chunk.choices?.[0]?.delta?.content;
        if (typeof content === 'string' && content.length > 0) {
          streamedChars += content.length;
          // Mark before yielding: once a token is downstream, retrying would duplicate it.
          if (progress) progress.yielded = true;
          yield content;
        }
      }

      // Spec 16 / F5: a "successful" stream that produced no content at all is a
      // reasoning-token exhaustion, not a parse problem. Make it visible.
      if (streamedChars === 0) {
        const completionTokens = lastUsage?.completionTokens ?? 0;
        const reasoningTokens = lastUsage?.reasoningTokens ?? 0;
        this.logger.error(
          `ALERT reasoning_token_exhaustion (stream) model=${model} ` +
            `maxCompletionTokens=${completionBudget} reasoningCap=${reasoningCap ?? 'none'} ` +
            `completionTokens=${completionTokens} reasoningTokens=${reasoningTokens}`,
        );
      }
    } catch (error: unknown) {
      const status = this.extractStatus(error);
      const detail = error instanceof Error ? error.message : 'OpenRouter request failed';
      if (status === 429 && telemetry?.modelTier) {
        this.modelRouter.markRateLimited(telemetry.modelTier as LLMTier);
      }
      this.logger.warn(`OpenRouter chat failed (${status}): ${detail}`);
      throw new LlmRequestError(status, detail);
    }
  }

  /**
   * Send a chat completion. If the model rejects disabled reasoning, retry once with effort=low.
   */
  private async requestChatCompletion(
    client: OpenRouterClient,
    opts: {
      model: string;
      messages: ChatMessage[];
      temperature: number;
      maxCompletionTokens: number;
      reasoningEffort: ReasoningEffort;
      reasoningMaxTokens?: number;
      responseFormat: 'json_object' | 'text';
    },
  ): Promise<ChatCompletionResult> {
    try {
      return await this.sendChatCompletionOnce(client, opts);
    } catch (error: unknown) {
      const status = this.extractStatus(error);
      if (opts.reasoningEffort !== 'low' && isReasoningMandatoryError(error, status)) {
        this.logger.warn(
          `OpenRouter rejected reasoning.effort=${opts.reasoningEffort} — retrying with effort=low`,
        );
        return await this.sendChatCompletionOnce(client, {
          ...opts,
          reasoningEffort: 'low',
        });
      }
      if (this.isTransientNetworkError(error)) {
        this.logger.warn(
          `OpenRouter transient network error (${this.describeNetworkError(error)}) — retrying once`,
        );
        return await this.sendChatCompletionOnce(client, opts);
      }
      throw error;
    }
  }

  private async sendChatCompletionOnce(
    client: OpenRouterClient,
    opts: {
      model: string;
      messages: ChatMessage[];
      temperature: number;
      maxCompletionTokens: number;
      reasoningEffort: ReasoningEffort;
      reasoningMaxTokens?: number;
      responseFormat: 'json_object' | 'text';
    },
  ): Promise<ChatCompletionResult> {
    const reasoning: { effort: ReasoningEffort; max_tokens?: number } = {
      effort: opts.reasoningEffort,
    };
    if (typeof opts.reasoningMaxTokens === 'number' && opts.reasoningMaxTokens > 0) {
      // OpenRouter / some models accept max_tokens on reasoning to leave room for output.
      reasoning.max_tokens = opts.reasoningMaxTokens;
    }

    const response = await client.chat.send({
      chatRequest: {
        model: opts.model,
        messages: opts.messages,
        ...(opts.responseFormat === 'json_object'
          ? { responseFormat: { type: 'json_object' } }
          : {}),
        stream: false,
        temperature: opts.temperature,
        maxCompletionTokens: opts.maxCompletionTokens,
        // SDK typings may omit max_tokens on reasoning — cast for providers that support it.
        reasoning: reasoning as { effort: ReasoningEffort },
      },
    });

    return response as ChatCompletionResult;
  }

  private extractCompletionText(response: ChatCompletionResult): string {
    const message = response.choices?.[0]?.message;
    const fromContent = extractChatContent(message?.content);
    if (fromContent.trim()) {
      return fromContent;
    }
    if (typeof message?.refusal === 'string' && message.refusal.trim()) {
      return message.refusal;
    }
    return '';
  }

  private logEmptyCompletion(
    model: string,
    response: ChatCompletionResult,
    attempt: string,
  ): void {
    const finishReason = response.choices?.[0]?.finishReason ?? 'unknown';
    const usage = response.usage;
    const completionTokens = usage?.completionTokens ?? 0;
    const reasoningTokens = usage?.completionTokensDetails?.reasoningTokens ?? 0;
    this.logger.warn(
      `OpenRouter empty content on ${attempt} (model=${model}, finishReason=${finishReason}, ` +
        `completionTokens=${completionTokens}, reasoningTokens=${reasoningTokens})`,
    );
    // Spec 16: full budget spent on reasoning with zero output — alertable production signal.
    if (completionTokens > 0 && completionTokens === reasoningTokens) {
      this.logger.error(
        `ALERT reasoning_token_exhaustion model=${model} attempt=${attempt} ` +
          `finishReason=${finishReason} completionTokens=${completionTokens} reasoningTokens=${reasoningTokens}`,
      );
    }
  }

  private extractStatus(error: unknown): number {
    if (this.isTransientNetworkError(error)) {
      return 503;
    }
    if (error && typeof error === 'object') {
      const record = error as Record<string, unknown>;
      if (typeof record.statusCode === 'number') {
        return record.statusCode;
      }
      if (typeof record.status === 'number') {
        return record.status;
      }
    }
    return 502;
  }

  private isTransientNetworkError(error: unknown): boolean {
    if (!error || typeof error !== 'object') return false;
    const err = error as Error & { cause?: { code?: string }; code?: string };
    const message = String(err.message ?? '').toLowerCase();
    const code = err.code ?? err.cause?.code;
    return (
      message === 'terminated' ||
      message.includes('econnreset') ||
      message.includes('fetch failed') ||
      message.includes('socket hang up') ||
      code === 'ECONNRESET' ||
      code === 'ECONNREFUSED' ||
      code === 'ETIMEDOUT' ||
      code === 'UND_ERR_SOCKET'
    );
  }

  private describeNetworkError(error: unknown): string {
    if (!(error instanceof Error)) return String(error);
    const cause = (error as Error & { cause?: { code?: string } }).cause;
    return cause?.code ? `${error.message} (${cause.code})` : error.message;
  }

  private normalizeUsage(usage: Record<string, unknown> | undefined): LlmUsage | undefined {
    if (!usage) {
      return undefined;
    }

    const completionDetails = usage.completionTokensDetails as
      | Record<string, unknown>
      | undefined;

    return {
      promptTokens: this.numberValue(usage.promptTokens) ?? this.numberValue(usage.prompt_tokens),
      completionTokens:
        this.numberValue(usage.completionTokens) ?? this.numberValue(usage.completion_tokens),
      totalTokens: this.numberValue(usage.totalTokens) ?? this.numberValue(usage.total_tokens),
      reasoningTokens:
        this.numberValue(completionDetails?.reasoningTokens) ??
        this.numberValue(completionDetails?.reasoning_tokens),
    };
  }

  private numberValue(value: unknown): number | undefined {
    return typeof value === 'number' ? value : undefined;
  }
}
