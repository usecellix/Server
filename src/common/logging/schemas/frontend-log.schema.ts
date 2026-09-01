import { TELEMETRY_CATEGORIES } from '../dto/frontend-log-batch.dto';
import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';
import { LOG_TTL_SECONDS } from './request-log.schema';

export type FrontendLogDocument = HydratedDocument<FrontendLog>;

export type FrontendLogLevel = 'error' | 'warn' | 'info' | 'action';
/**
 * Derived from the DTO's allowlist rather than restated — TASKS.md #159.
 *
 * This list previously existed FOUR times (client union, request DTO, this
 * file-logger type, and the Mongo schema), independently maintained. Adding
 * `'verify'` in one place 400'd every telemetry batch containing it, and since
 * the DTO rejects a whole batch, `accept.success` was silently lost alongside
 * it. Deriving keeps the compiler honest; `test/telemetry-category-parity.spec.ts`
 * covers the one copy that cannot be derived (the client, across the repo
 * boundary).
 */
export type FrontendLogCategory = (typeof TELEMETRY_CATEGORIES)[number];

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
