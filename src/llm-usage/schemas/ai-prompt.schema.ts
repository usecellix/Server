import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, Types } from 'mongoose';

export type AiPromptDocument = HydratedDocument<AiPrompt>;

/**
 * One row per user prompt, with totals kept current by `$inc` on every LLM
 * call so the admin list never has to aggregate `llm_calls` to render. A
 * stepwise build's /continue waves roll up into the prompt that started it.
 */
@Schema({ collection: 'ai_prompts', versionKey: false })
export class AiPrompt {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, unique: true })
  promptId!: string;

  @Prop({ type: String, index: true })
  userId?: string;

  @Prop({ type: String, index: true })
  conversationId?: string;

  @Prop({ type: String })
  workbookId?: string;

  @Prop({ type: String })
  prompt?: string;

  @Prop({ type: String })
  mode?: string;

  @Prop({ type: String })
  route?: string;

  @Prop({ type: Number })
  tier?: number;

  @Prop({ type: Date, required: true, default: () => new Date(), index: true })
  createdAt!: Date;

  @Prop({ type: Date, required: true, default: () => new Date() })
  lastActivityAt!: Date;

  /** HTTP requests that worked on this prompt: the first one plus each /continue. */
  @Prop({ type: Number, default: 0 })
  requestCount!: number;

  @Prop({ type: Number, default: 0 })
  llmCalls!: number;

  @Prop({ type: Number, default: 0 })
  failedCalls!: number;

  @Prop({ type: Number, default: 0 })
  promptTokens!: number;

  @Prop({ type: Number, default: 0 })
  completionTokens!: number;

  @Prop({ type: Number, default: 0 })
  reasoningTokens!: number;

  @Prop({ type: Number, default: 0 })
  cachedTokens!: number;

  @Prop({ type: Number, default: 0 })
  totalTokens!: number;

  @Prop({ type: Number, default: 0 })
  costUsd!: number;

  /** Summed per-call latency — model time, not wall-clock (calls run in parallel). */
  @Prop({ type: Number, default: 0 })
  llmLatencyMs!: number;

  /** Wall-clock time spent inside HTTP requests for this prompt. */
  @Prop({ type: Number, default: 0 })
  requestDurationMs!: number;

  @Prop({ type: [String], default: [] })
  models!: string[];

  @Prop({ type: String, enum: ['running', 'ok', 'error'], default: 'running' })
  lastOutcome!: 'running' | 'ok' | 'error';

  @Prop({ type: String })
  lastError?: string;
}

export const AiPromptSchema = SchemaFactory.createForClass(AiPrompt);
AiPromptSchema.index({ userId: 1, createdAt: -1 });
AiPromptSchema.index({ costUsd: -1 });
