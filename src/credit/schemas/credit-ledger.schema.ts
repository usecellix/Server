import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, Types } from 'mongoose';

export type CreditLedgerEntryDocument = HydratedDocument<CreditLedgerEntry>;

/**
 * Append-only. Never updated, never deleted — CREDIT_SYSTEM.md CD-9, the
 * same durable-record convention as ChangeSet/AuditLog. No TTL: this is a
 * financial record, not working memory (CREDIT_SYSTEM_SCHEMA.md §3).
 */
@Schema({ collection: 'credit_ledger', versionKey: false })
export class CreditLedgerEntry {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, index: true })
  billingEntityId!: string;

  /** WHO spent it, for pooled org accounts. Absent for user-billed accounts. */
  @Prop({ type: String, index: true })
  seatUserId?: string;

  @Prop({ type: String, required: true, enum: ['grant', 'purchase', 'debit', 'one_time_grant'] })
  entryType!: 'grant' | 'purchase' | 'debit' | 'one_time_grant';

  /** Positive for grant/purchase, negative for debit. */
  @Prop({ type: Number, required: true })
  amount!: number;

  @Prop({ type: String, required: true, enum: ['planCredits', 'purchasedCredits', 'oneTimeCredits'] })
  bucket!: 'planCredits' | 'purchasedCredits' | 'oneTimeCredits';

  /** Set for entryType === 'debit'. Matches a CreditActionType catalog key. */
  @Prop({ type: String })
  actionType?: string;

  @Prop({ type: String })
  conversationId?: string;

  @Prop({ type: String })
  changeSetId?: string;

  /** Idempotency key for grant/purchase entries originating from a Stripe webhook. */
  @Prop({ type: String, index: true, sparse: true, unique: true })
  stripeEventId?: string;

  @Prop({ type: Date, required: true, default: () => new Date(), index: true })
  createdAt!: Date;
}

export const CreditLedgerEntrySchema = SchemaFactory.createForClass(CreditLedgerEntry);
CreditLedgerEntrySchema.index({ billingEntityId: 1, createdAt: -1 });
CreditLedgerEntrySchema.index({ seatUserId: 1, createdAt: -1 });
