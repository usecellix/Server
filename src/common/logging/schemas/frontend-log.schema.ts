import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';
import { LOG_TTL_SECONDS } from './request-log.schema';

export type FrontendLogDocument = HydratedDocument<FrontendLog>;

export type FrontendLogLevel = 'error' | 'warn' | 'info' | 'action';
export type FrontendLogCategory =
  | 'console'
  | 'preview'
  | 'accept'
  | 'reject'
  | 'apply'
  | 'sse'
  | 'navigation'
  | 'other';

@Schema({
  collection: 'frontend_logs',
  versionKey: false,
})
export class FrontendLog {
  _id!: Types.ObjectId;

  @Prop({ type: Date, required: true, index: true })
  ts!: Date;

  @Prop({ type: String, required: true, index: true })
  level!: FrontendLogLevel;

  @Prop({ type: String, required: true, index: true })
  category!: FrontendLogCategory;

  @Prop({ type: String, required: true, index: true })
  event!: string;

  @Prop({ type: String, required: true })
  message!: string;

  @Prop({ type: String, index: true })
  conversationId?: string;

  @Prop({ type: String, index: true })
  changeSetId?: string;

  @Prop({ type: String, index: true })
  sessionId?: string;

  @Prop({ type: String })
  workbookKey?: string;

  @Prop({ type: String })
  userAgent?: string;

  @Prop({ type: String })
  pageUrl?: string;

  @Prop({ type: SchemaTypes.Mixed })
  details?: unknown;
}

export const FrontendLogSchema = SchemaFactory.createForClass(FrontendLog);
FrontendLogSchema.index({ ts: 1 }, { expireAfterSeconds: LOG_TTL_SECONDS });
