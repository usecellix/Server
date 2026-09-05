import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, Types } from 'mongoose';

export type CreditAccountDocument = HydratedDocument<CreditAccount>;

/**
 * One per billing entity — a user (Free/Solo) or an org (Firm/Enterprise,
 * pooled across seats). CREDIT_SYSTEM.md CD-7, CD-8.
 */
@Schema({ collection: 'credit_accounts', versionKey: false, timestamps: true })
export class CreditAccount {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, enum: ['user', 'org'] })
  billingEntityType!: 'user' | 'org';

  @Prop({ type: String, required: true, unique: true, index: true })
  billingEntityId!: string;

  @Prop({ type: String, required: true, enum: ['free', 'solo', 'firm', 'enterprise'], index: true })
  planTier!: 'free' | 'solo' | 'firm' | 'enterprise';

  /** Resets to the plan's monthly allotment each cycle. Does not roll over. */
  @Prop({ type: Number, required: true, default: 0 })
  planCredits!: number;

  /** Top-up pack credits. Persist indefinitely, never expire. */
  @Prop({ type: Number, required: true, default: 0 })
  purchasedCredits!: number;

  /** Free tier's 30-credit grant. Issued once at signup, never reset. */
  @Prop({ type: Number, required: true, default: 0 })
  oneTimeCredits!: number;

  @Prop({ type: Date })
  currentPeriodStart?: Date;

  @Prop({ type: Date })
  currentPeriodEnd?: Date;

  /** Only set when billingEntityType === 'org'. */
  @Prop({ type: [String] })
  seatUserIds?: string[];

  createdAt!: Date;
  updatedAt!: Date;
}

export const CreditAccountSchema = SchemaFactory.createForClass(CreditAccount);
CreditAccountSchema.index({ billingEntityType: 1, planTier: 1 });
