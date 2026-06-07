import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';

export type AuditEntryDocument = HydratedDocument<AuditEntry>;

@Schema({
  collection: 'audit_entries',
  timestamps: true,
  versionKey: false,
})
export class AuditEntry {
  @Prop({ type: String, required: true, index: true })
  requestId!: string;

  @Prop({ type: String, required: true, index: true })
  processName!: string;

  @Prop({ type: String, required: true })
  action!: string;

  @Prop({ type: String })
  userId?: string;

  @Prop({ type: Number, min: 0, max: 1 })
  confidence?: number;

  @Prop({ type: SchemaTypes.Mixed })
  payload?: unknown;

  @Prop({ type: SchemaTypes.Mixed })
  result?: unknown;

  _id!: Types.ObjectId;
  createdAt!: Date;
  updatedAt!: Date;
}

export const AuditEntrySchema = SchemaFactory.createForClass(AuditEntry);
