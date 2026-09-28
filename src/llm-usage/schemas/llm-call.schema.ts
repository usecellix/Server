import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, Types } from 'mongoose';

export type LlmCallDocument = HydratedDocument<LlmCall>;

/**
 * One row per network attempt to the LLM provider — retries included, since a
 * retried call is billed like any other. No TTL: this is the cost record the
 * admin dashboard and any future billing reconciliation read from.
 */
@Schema({ collection: 'llm_calls', versionKey: false })
export class LlmCall {
  _id!: Types.ObjectId;

  /** Null when the call happened outside any user request (e.g. web chat, startup). */
  @Prop({ type: String, default: null })
  promptId!: string | null;

  @Prop({ type: String })
  userId?: string;

  @Prop({ type: String })
  conversationId?: string;

  @Prop({ type: Date, required: true, default: () => new Date() })
  ts!: Date;

  /** Model we asked for. */
  @Prop({ type: String, required: true })
  model!: string;

  /** Model the provider says actually served the call (routing can differ). */
  @Prop({ type: String })
  servedModel?: string;

  @Prop({ type: String })
  generationId?: string;

  /** Which part of the pipeline made the call (planner, executor, router, …). */
  @Prop({ type: String, required: true })
  caller!: string;

  /** 1-based attempt within one logical call; >1 means a retry. */
  @Prop({ type: Number, required: true, default: 1 })
  attempt!: number;

  @Prop({ type: Boolean, required: true, default: false })
  streaming!: boolean;

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

  /** True when the provider returned no cost and we priced it from MODEL_CONFIGS. */
  @Prop({ type: Boolean, default: false })
  costEstimated!: boolean;

  @Prop({ type: Number, required: true })
  latencyMs!: number;

  @Prop({ type: Boolean, required: true })
  success!: boolean;

  @Prop({ type: String })
  finishReason?: string;

  @Prop({ type: Number })
  errorStatus?: number;

  @Prop({ type: String })
  errorMessage?: string;
}

export const LlmCallSchema = SchemaFactory.createForClass(LlmCall);
LlmCallSchema.index({ promptId: 1, ts: 1 });
LlmCallSchema.index({ ts: -1 });
LlmCallSchema.index({ userId: 1, ts: -1 });
LlmCallSchema.index({ model: 1, ts: -1 });
