import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, Types } from 'mongoose';

export type CheckpointDocument = HydratedDocument<Checkpoint>;

export type CheckpointTrigger = 'auto' | 'manual';
export type CheckpointStatus = 'active' | 'restored';

/**
 * PRD M5.2 / TASKS.md #26. A checkpoint is a marker pointing at the last
 * `change_sets` document already applied at capture time — restore is a
 * service-layer algorithm (CheckpointService.restore) that walks and inverts
 * every change set applied after this anchor, not stored state on this
 * document itself (DATABASE_SCHEMA.md §6.2).
 */
@Schema({
  collection: 'checkpoints',
  versionKey: false,
})
export class Checkpoint {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, unique: true, index: true })
  checkpointId!: string;

  /** Durable per-workbook identity — restore is scoped by this, never by conversationId. */
  @Prop({ type: String, required: true, index: true })
  workbookId!: string;

  /** The session that created it — display-only, not load-bearing for restore. */
  @Prop({ type: String, required: true, index: true })
  conversationId!: string;

  @Prop({ type: String, required: true })
  label!: string;

  @Prop({ type: String, required: true, enum: ['auto', 'manual'] })
  trigger!: CheckpointTrigger;

  /**
   * The last applied change_sets document at capture time. '' means "no prior
   * applied change sets" — a real, meaningful value, not a missing field, so
   * this isn't `required: true` (Mongoose's String required check rejects '').
   */
  @Prop({ type: String, required: false, default: '', index: true })
  anchorChangeSetId!: string;

  @Prop({ type: Date, required: true, default: () => new Date() })
  createdAt!: Date;

  @Prop({ type: String, required: true, enum: ['active', 'restored'], default: 'active' })
  status!: CheckpointStatus;

  @Prop({ type: Date })
  restoredAt?: Date;
}

export const CheckpointSchema = SchemaFactory.createForClass(Checkpoint);
// "List checkpoints for this workbook, newest first" — the actual query shape #31's panel needs.
CheckpointSchema.index({ workbookId: 1, createdAt: -1 });
