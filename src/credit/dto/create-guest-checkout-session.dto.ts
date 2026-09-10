import { IsEmail, IsIn } from 'class-validator';

/**
 * Body for `POST /billing/checkout/subscribe-by-email` — the marketing
 * site's pre-signup checkout flow, where the visitor has no product session
 * yet. See razorpay-checkout.service.ts's docblock on normalizeGuestEmail for
 * why the email itself becomes billingEntityId.
 */
export class CreateGuestCheckoutSessionDto {
  @IsEmail()
  email!: string;

  @IsIn(['solo', 'firm', 'beta'])
  planTier!: 'solo' | 'firm' | 'beta';
}
