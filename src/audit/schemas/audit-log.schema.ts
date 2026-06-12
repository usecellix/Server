import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';

export type AuditLogDocument = HydratedDocument<AuditLog>;

@Schema({
  collection: 'audit_logs',
  versionKey: false,
})
export class AuditLog {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, index: true })
  traceId!: string;

  @Prop({ type: Date, required: true, default: () => new Date(), index: true })
  timestamp!: Date;

  @Prop({ type: String, required: true })
  llmModel!: string;

  @Prop({ type: String, required: true, enum: ['low', 'medium', 'high'] })
  tier!: string;

  @Prop({ type: String, required: true })
  intent!: string;

  @Prop({ type: Number, required: true })
  promptTokens!: number;

  @Prop({ type: Number, required: true })
  completionTokens!: number;

  @Prop({ type: Number, required: true })
  totalTokens!: number;

  @Prop({ type: Number, required: true })
  estimatedCostUsd!: number;

  @Prop({ type: Number, required: true })
  latencyMs!: number;

  @Prop({ type: Boolean, required: true })
  success!: boolean;

  @Prop({ type: String })
  errorCode?: string;

  @Prop({ type: Number })
  actionsCount?: number;

  @Prop({ type: SchemaTypes.Mixed })
  rawUsage?: Record<string, unknown>;
}

export const AuditLogSchema = SchemaFactory.createForClass(AuditLog);
