import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../../config/app-config.service';
import { LLMTier } from '../../types/cellix.types';
import { LlmRequestError } from '../errors/llm-request.error';
import { ModelRouter } from '../llm/model-router';

export type OpenRouterChatMessage = {
  role: 'system' | 'user' | 'assistant';
  content: string;
};

export type LlmUsage = {
  promptTokens?: number;
  completionTokens?: number;
  totalTokens?: number;
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
    temperature?: number;
    maxTokens?: number;
  }): Promise<string> {
    const apiKey = this.config.openRouterApiKey;
    if (!apiKey) {
      throw new LlmRequestError(503, 'OpenRouter not configured');
    }

    try {
      const { OpenRouter } = await import('@openrouter/sdk');
      const client = new OpenRouter({
        apiKey,
        httpReferer: this.config.openRouterHttpReferer,
        appTitle: 'Cellix',
      });

      const response = await client.chat.send({
        chatRequest: {
          model: opts.model ?? this.config.openRouterModelMedium,
          messages: [
            { role: 'system', content: opts.systemPrompt },
            { role: 'user', content: opts.userMessage },
          ],
          stream: false,
          temperature: opts.temperature ?? 0.2,
          maxTokens: opts.maxTokens ?? 1500,
        },
      });

      const content = response.choices?.[0]?.message?.content;
      return typeof content === 'string' ? content : '';
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

      const response = await client.chat.send({
        chatRequest: {
          model: this.config.openRouterModelLow,
          messages: [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userMessage },
          ],
          stream: false,
          temperature: 0,
          maxTokens: 256,
        },
      });

      const content = response.choices?.[0]?.message?.content;
      return typeof content === 'string' ? content : '';
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
          maxTokens,
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

    return {
      promptTokens: this.numberValue(usage.promptTokens) ?? this.numberValue(usage.prompt_tokens),
      completionTokens:
        this.numberValue(usage.completionTokens) ?? this.numberValue(usage.completion_tokens),
      totalTokens: this.numberValue(usage.totalTokens) ?? this.numberValue(usage.total_tokens),
    };
  }

  private numberValue(value: unknown): number | undefined {
    return typeof value === 'number' ? value : undefined;
  }
}
