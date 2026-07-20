import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';
import { LOG_TTL_SECONDS } from './request-log.schema';

export type PlannerLogDocument = HydratedDocument<PlannerLog>;

@Schema({
  collection: 'planner_logs',
  versionKey: false,
})
export class PlannerLog {
  _id!: Types.ObjectId;

  @Prop({ type: Date, required: true })
  ts!: Date;

  @Prop({ type: String, required: true, index: true })
  correlationId!: string;

  @Prop({ type: String, required: true })
  model!: string;

  @Prop({ type: Number, required: true })
  durationMs!: number;

  @Prop({ type: Boolean, required: true, index: true })
  success!: boolean;

  @Prop({ type: String })
  error?: string;

  @Prop({ type: SchemaTypes.Mixed, required: true })
  input!: Record<string, unknown>;

  @Prop({ type: SchemaTypes.Mixed, required: true })
  output!: Record<string, unknown>;
}

export const PlannerLogSchema = SchemaFactory.createForClass(PlannerLog);
PlannerLogSchema.index({ ts: 1 }, { expireAfterSeconds: LOG_TTL_SECONDS });
