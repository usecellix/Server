import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../../config/app-config.service';
import { LLMTier } from '../../types/cellix.types';
import { LlmRequestError } from '../errors/llm-request.error';
import { ModelRouter } from '../llm/model-router';
import { extractChatContent } from '../utils/extract-chat-content.util';

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
    responseFormat?: 'json_object' | 'text';
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
    const minCompletionBudget = budget <= 512 ? budget : Math.max(budget, 4096);

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
        maxCompletionTokens: minCompletionBudget,
        reasoningEffort: opts.reasoningEffort ?? 'low',
        responseFormat: opts.responseFormat ?? 'json_object',
      });

      let text = this.extractCompletionText(response);
      if (!text.trim()) {
        this.logEmptyCompletion(model, response, 'first attempt');
        response = await this.requestChatCompletion(client, {
          model,
          messages: baseMessages,
          temperature: opts.temperature ?? 0.2,
          maxCompletionTokens: Math.max(minCompletionBudget, 8192),
          reasoningEffort: 'none',
          responseFormat: opts.responseFormat ?? 'json_object',
        });
        text = this.extractCompletionText(response);
        if (!text.trim()) {
          this.logEmptyCompletion(model, response, 'retry with reasoning.effort=none');
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

  async *streamChat(
    messages: OpenRouterChatMessage[],
    telemetry?: LlmCallTelemetry,
    model = this.config.openRouterModelMedium,
    maxTokens = 4096,
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

      const stream = await client.chat.send({
        chatRequest: {
          model,
          messages,
          stream: true,
          streamOptions: { includeUsage: true },
          temperature: 0.25,
          maxCompletionTokens: maxTokens,
        },
      });

      for await (const chunk of stream) {
        const usage = this.normalizeUsage((chunk as { usage?: Record<string, unknown> }).usage);
        if (usage && telemetry) {
          telemetry.usage = usage;
        }

        const content = chunk.choices?.[0]?.delta?.content;
        if (typeof content === 'string' && content.length > 0) {
          yield content;
        }
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

  private async requestChatCompletion(
    client: InstanceType<Awaited<typeof import('@openrouter/sdk')>['OpenRouter']>,
    opts: {
      model: string;
      messages: ChatMessage[];
      temperature: number;
      maxCompletionTokens: number;
      reasoningEffort: ReasoningEffort;
      responseFormat: 'json_object' | 'text';
    },
  ): Promise<ChatCompletionResult> {
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
        reasoning: { effort: opts.reasoningEffort },
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
    const reasoningTokens = usage?.completionTokensDetails?.reasoningTokens ?? 0;
    this.logger.warn(
      `OpenRouter empty content on ${attempt} (model=${model}, finishReason=${finishReason}, ` +
        `completionTokens=${usage?.completionTokens ?? 0}, reasoningTokens=${reasoningTokens})`,
    );
  }

  private extractStatus(error: unknown): number {
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
