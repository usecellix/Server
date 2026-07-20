import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';

/** Auto-delete log documents 3 days after `ts`. */
export const LOG_TTL_SECONDS = 3 * 24 * 60 * 60;

export type RequestLogDocument = HydratedDocument<RequestLog>;

@Schema({
  collection: 'request_logs',
  versionKey: false,
})
export class RequestLog {
  _id!: Types.ObjectId;

  @Prop({ type: Date, required: true })
  ts!: Date;

  @Prop({ type: String, required: true })
  method!: string;

  @Prop({ type: String, required: true, index: true })
  url!: string;

  @Prop({ type: Number, required: true })
  statusCode!: number;

  @Prop({ type: Number, required: true })
  responseTimeMs!: number;

  @Prop({ type: String, index: true })
  reqId?: string;

  @Prop({ type: String, index: true })
  traceId?: string;

  @Prop({ type: String })
  message?: string;

  @Prop({ type: SchemaTypes.Mixed })
  response?: unknown;
}

export const RequestLogSchema = SchemaFactory.createForClass(RequestLog);
RequestLogSchema.index({ ts: 1 }, { expireAfterSeconds: LOG_TTL_SECONDS });
