import { Injectable, Logger } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { currentLlmUsageContext, type LlmUsageContext } from './llm-usage.context';
import { AiPrompt } from './schemas/ai-prompt.schema';
import { LlmCall } from './schemas/llm-call.schema';

const PROMPT_TEXT_MAX = 4000;
const ERROR_TEXT_MAX = 500;

export interface RecordLlmCallInput {
  model: string;
  servedModel?: string;
  generationId?: string;
  caller: string;
  attempt: number;
  streaming: boolean;
  promptTokens?: number;
  completionTokens?: number;
  reasoningTokens?: number;
  cachedTokens?: number;
  totalTokens?: number;
  costUsd: number;
  costEstimated: boolean;
  latencyMs: number;
  success: boolean;
  finishReason?: string | null;
  errorStatus?: number;
  errorMessage?: string;
}

export interface BeginPromptInput {
  prompt?: string;
  mode?: string;
  workbookId?: string;
}

/**
 * Writes are fire-and-forget: accounting must never slow down or fail a user
 * request. A failed write is logged, not thrown.
 */
@Injectable()
export class LlmUsageService {
  private readonly logger = new Logger(LlmUsageService.name);

  constructor(
    @InjectModel(AiPrompt.name) private readonly promptModel: Model<AiPrompt>,
    @InjectModel(LlmCall.name) private readonly callModel: Model<LlmCall>,
  ) {}

  beginPrompt(context: LlmUsageContext, input: BeginPromptInput): void {
    void this.upsertPrompt(context.promptId, {
      $setOnInsert: { createdAt: new Date() },
      $set: {
        lastActivityAt: new Date(),
        lastOutcome: 'running',
        ...(context.userId ? { userId: context.userId } : {}),
        ...(context.conversationId ? { conversationId: context.conversationId } : {}),
        ...(input.prompt ? { prompt: input.prompt.slice(0, PROMPT_TEXT_MAX) } : {}),
        ...(input.mode ? { mode: input.mode } : {}),
        ...(input.workbookId ? { workbookId: input.workbookId } : {}),
      },
    });
  }

  endRequest(context: LlmUsageContext, outcome: { durationMs: number; error?: string }): void {
    const error = outcome.error ?? context.error;
    void this.upsertPrompt(context.promptId, {
      $setOnInsert: { createdAt: new Date() },
      $inc: { requestCount: 1, requestDurationMs: Math.max(0, Math.round(outcome.durationMs)) },
      $set: {
        lastActivityAt: new Date(),
        lastOutcome: error ? 'error' : 'ok',
        ...(error ? { lastError: error.slice(0, ERROR_TEXT_MAX) } : {}),
        ...contextFields(context),
      },
    });
  }

  recordCall(input: RecordLlmCallInput): void {
    const context = currentLlmUsageContext();
    const promptTokens = input.promptTokens ?? 0;
    const completionTokens = input.completionTokens ?? 0;
    const totalTokens = input.totalTokens ?? promptTokens + completionTokens;

    const promptId = context?.promptId || null;
    void this.callModel
      .create({
        promptId,
        userId: context?.userId,
        conversationId: context?.conversationId,
        ts: new Date(),
        model: input.model,
        servedModel: input.servedModel,
        generationId: input.generationId,
        caller: input.caller,
        attempt: input.attempt,
        streaming: input.streaming,
        promptTokens,
        completionTokens,
        reasoningTokens: input.reasoningTokens ?? 0,
        cachedTokens: input.cachedTokens ?? 0,
        totalTokens,
        costUsd: input.costUsd,
        costEstimated: input.costEstimated,
        latencyMs: Math.round(input.latencyMs),
        success: input.success,
        ...(input.finishReason ? { finishReason: input.finishReason } : {}),
        ...(input.errorStatus !== undefined ? { errorStatus: input.errorStatus } : {}),
        ...(input.errorMessage ? { errorMessage: input.errorMessage.slice(0, ERROR_TEXT_MAX) } : {}),
      })
      .catch((err: unknown) => this.warn('record llm call', err));

    if (!context || !promptId) return;
    void this.upsertPrompt(promptId, {
      $setOnInsert: { createdAt: new Date() },
      $inc: {
        llmCalls: 1,
        failedCalls: input.success ? 0 : 1,
        promptTokens,
        completionTokens,
        reasoningTokens: input.reasoningTokens ?? 0,
        cachedTokens: input.cachedTokens ?? 0,
        totalTokens,
        costUsd: input.costUsd,
        llmLatencyMs: Math.round(input.latencyMs),
      },
      $addToSet: { models: input.servedModel ?? input.model },
      $set: { lastActivityAt: new Date(), ...contextFields(context) },
    });
  }

  private async upsertPrompt(promptId: string, update: Record<string, unknown>): Promise<void> {
    if (!promptId) return;
    // Two concurrent upserts for a brand-new promptId can race on the unique
    // index; the loser's retry becomes a plain update of the winner's row.
    for (let attempt = 0; attempt < 2; attempt += 1) {
      try {
        await this.promptModel.updateOne({ promptId }, update, { upsert: true }).exec();
        return;
      } catch (err: unknown) {
        if (attempt === 0 && (err as { code?: number })?.code === 11000) continue;
        this.warn('upsert prompt', err);
        return;
      }
    }
  }

  private warn(what: string, err: unknown): void {
    const msg = err instanceof Error ? err.message : String(err);
    this.logger.warn(`Failed to ${what}: ${msg}`);
  }
}

function contextFields(context: LlmUsageContext): Record<string, unknown> {
  return {
    ...(context.userId ? { userId: context.userId } : {}),
    ...(context.conversationId ? { conversationId: context.conversationId } : {}),
    ...(context.route ? { route: context.route } : {}),
    ...(context.tier !== undefined ? { tier: context.tier } : {}),
  };
}
