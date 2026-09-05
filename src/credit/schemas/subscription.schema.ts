import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, Types } from 'mongoose';

export type SubscriptionDocument = HydratedDocument<Subscription>;

/**
 * Stripe-synced, read mostly by webhook handlers. CREDIT_SYSTEM_SCHEMA.md
 * §4. Free has no subscription row; Enterprise custom-plan handling is an
 * open question (CREDIT_SYSTEM.md §8 Q2) — not modeled here yet.
 */
@Schema({ collection: 'subscriptions', versionKey: false, timestamps: true })
export class Subscription {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, index: true })
  billingEntityId!: string;

  @Prop({ type: String, required: true })
  stripeCustomerId!: string;

  @Prop({ type: String, required: true, unique: true, index: true })
  stripeSubscriptionId!: string;

  @Prop({ type: String, required: true, enum: ['active', 'past_due', 'canceled', 'incomplete'] })
  status!: 'active' | 'past_due' | 'canceled' | 'incomplete';

  @Prop({ type: String, required: true, enum: ['solo', 'firm'] })
  planTier!: 'solo' | 'firm';

  @Prop({ type: Date, required: true })
  currentPeriodEnd!: Date;

  @Prop({ type: Boolean, required: true, default: false })
  cancelAtPeriodEnd!: boolean;

  createdAt!: Date;
  updatedAt!: Date;
}

export const SubscriptionSchema = SchemaFactory.createForClass(Subscription);

export type ProcessedStripeEventDocument = HydratedDocument<ProcessedStripeEvent>;

/**
 * Webhook idempotency log, separate from credit_ledger.stripeEventId
 * because a subscription-status-only webhook (e.g.
 * customer.subscription.updated with no credit grant) never produces a
 * ledger row. CREDIT_SYSTEM_SCHEMA.md §4.
 */
@Schema({ collection: 'processed_stripe_events', versionKey: false })
export class ProcessedStripeEvent {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, unique: true, index: true })
  stripeEventId!: string;

  @Prop({ type: Date, required: true, default: () => new Date() })
  processedAt!: Date;
}

export const ProcessedStripeEventSchema = SchemaFactory.createForClass(ProcessedStripeEvent);
