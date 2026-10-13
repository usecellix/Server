import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, Types } from 'mongoose';

export type TopupOrderDocument = HydratedDocument<TopupOrder>;

export type TopupOrderStatus = 'created' | 'paid' | 'expired' | 'cancelled';

/**
 * One row per top-up Payment Link we create. Lets a paid top-up be granted by
 * asking Razorpay about the link directly (RazorpayWebhookService.
 * reconcileTopups) instead of depending only on the `payment_link.paid`
 * webhook, which never arrives when Razorpay can't reach the backend (local
 * dev, a misconfigured dashboard webhook, an outage). Durable — a purchase
 * record, not working memory.
 */
@Schema({ collection: 'topup_orders', versionKey: false, timestamps: true })
export class TopupOrder {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, unique: true })
  paymentLinkId!: string;

  @Prop({ type: String, required: true, index: true })
  billingEntityId!: string;

  @Prop({ type: String, required: true, enum: ['small', 'medium', 'large'] })
  packId!: 'small' | 'medium' | 'large';

  @Prop({ type: Number, required: true })
  credits!: number;

  @Prop({ type: Number, required: true })
  priceInr!: number;

  @Prop({ type: String, required: true, enum: ['created', 'paid', 'expired', 'cancelled'], default: 'created' })
  status!: TopupOrderStatus;

  createdAt!: Date;
  updatedAt!: Date;
}

export const TopupOrderSchema = SchemaFactory.createForClass(TopupOrder);
TopupOrderSchema.index({ billingEntityId: 1, status: 1, createdAt: -1 });
