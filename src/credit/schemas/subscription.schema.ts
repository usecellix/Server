import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, Types } from 'mongoose';

export type SubscriptionDocument = HydratedDocument<Subscription>;

/**
 * Razorpay-synced, read mostly by webhook handlers. CREDIT_SYSTEM_SCHEMA.md
 * §4. Free has no subscription row; Enterprise custom-plan handling is an
 * open question (CREDIT_SYSTEM.md §8 Q2) — not modeled here yet.
 *
 * Migrated from Stripe (TASKS.md #181) to Razorpay (credit-system-v2 session)
 * — field names and the `status` enum are Razorpay-shaped, not Stripe's.
 * Razorpay's subscription status vocabulary does NOT map 1:1 onto Stripe's
 * (`active | past_due | canceled | incomplete`) — it has more granular
 * pre-activation states (`created`, `authenticated`, `pending`) and a
 * `halted` state Stripe has no equivalent for.
 */
@Schema({ collection: 'subscriptions', versionKey: false, timestamps: true })
export class Subscription {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, index: true })
  billingEntityId!: string;

  @Prop({ type: String })
  razorpayCustomerId?: string;

  @Prop({ type: String, required: true, unique: true, index: true })
  razorpaySubscriptionId!: string;

  @Prop({
    type: String,
    required: true,
    enum: ['created', 'authenticated', 'active', 'pending', 'halted', 'cancelled', 'completed', 'expired'],
  })
  status!: 'created' | 'authenticated' | 'active' | 'pending' | 'halted' | 'cancelled' | 'completed' | 'expired';

  @Prop({ type: String, required: true, enum: ['solo', 'firm', 'beta'] })
  planTier!: 'solo' | 'firm' | 'beta';

  @Prop({ type: Date })
  currentPeriodStart?: Date;

  @Prop({ type: Date, required: true })
  currentPeriodEnd!: Date;

  @Prop({ type: Boolean, required: true, default: false })
  cancelAtPeriodEnd!: boolean;

  createdAt!: Date;
  updatedAt!: Date;
}

export const SubscriptionSchema = SchemaFactory.createForClass(Subscription);

export type ProcessedRazorpayEventDocument = HydratedDocument<ProcessedRazorpayEvent>;

/**
 * Webhook idempotency log, separate from credit_ledger.paymentEventId
 * because a subscription-status-only webhook (e.g. subscription.updated with
 * no credit grant) never produces a ledger row. CREDIT_SYSTEM_SCHEMA.md §4.
 *
 * Razorpay's webhook payload has no single top-level `event.id` the way
 * Stripe's does — `paymentEventId` here is a composite key built by the
 * caller (RazorpayWebhookService), typically
 * `${event}:${entityId}:${paymentOrSubscriptionId}`, since that's the
 * smallest combination that's actually unique per real occurrence (a
 * `subscription.charged` event recurs every renewal for the SAME
 * subscription id, so the subscription id alone is not enough).
 */
@Schema({ collection: 'processed_razorpay_events', versionKey: false })
export class ProcessedRazorpayEvent {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, unique: true, index: true })
  paymentEventId!: string;

  @Prop({ type: Date, required: true, default: () => new Date() })
  processedAt!: Date;
}

export const ProcessedRazorpayEventSchema = SchemaFactory.createForClass(ProcessedRazorpayEvent);
