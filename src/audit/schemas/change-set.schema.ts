import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';

export type ChangeSetDocument = HydratedDocument<ChangeSet>;

@Schema({ _id: false })
export class CellSnapshotSchema {
  @Prop({ type: SchemaTypes.Mixed })
  value!: unknown;

  @Prop({ type: String, default: '' })
  formula!: string;

  @Prop({ type: String, default: 'General' })
  format!: string;
}

@Schema({ _id: false })
export class CellChangeSchema {
  @Prop({ type: String, required: true })
  cell!: string;

  @Prop({ type: String, required: true })
  sheet!: string;

  @Prop({ type: SchemaTypes.Mixed })
  before!: unknown;

  @Prop({ type: SchemaTypes.Mixed })
  after!: unknown;

  @Prop({ type: String })
  formula?: string;

  @Prop({ type: Boolean, required: true })
  isHardcoded!: boolean;

  @Prop({ type: [SchemaTypes.Mixed], required: false })
  sourceRefs?: Record<string, unknown>[];

  @Prop({ type: [SchemaTypes.Mixed], required: false })
  exceptionFlags?: Record<string, unknown>[];
}

@Schema({
  collection: 'change_sets',
  versionKey: false,
})
export class ChangeSet {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, unique: true, index: true })
  changeSetId!: string;

  @Prop({ type: String, required: true, index: true })
  conversationId!: string;

  @Prop({ type: String, required: true, index: true })
  traceId!: string;

  @Prop({ type: Date, required: true, default: () => new Date() })
  timestamp!: Date;

  @Prop({ type: String, required: true })
  prompt!: string;

  @Prop({ type: SchemaTypes.Mixed, default: {} })
  beforeState!: Record<string, CellSnapshotSchema>;

  @Prop({ type: [CellChangeSchema], default: [] })
  changes!: CellChangeSchema[];

  @Prop({ type: [SchemaTypes.Mixed], default: [] })
  actions!: Record<string, unknown>[];

  @Prop({
    type: String,
    required: true,
    enum: ['previewed', 'applied', 'reverted'],
    default: 'previewed',
  })
  status!: string;

  @Prop({ type: Date })
  appliedAt?: Date;

  @Prop({ type: Date })
  revertedAt?: Date;

  @Prop({ type: Number, required: false })
  provenanceConfidence?: number;
}

export const ChangeSetSchema = SchemaFactory.createForClass(ChangeSet);
