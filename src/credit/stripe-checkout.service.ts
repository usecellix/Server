import { BadRequestException, Injectable, ServiceUnavailableException } from '@nestjs/common';
import Stripe from 'stripe';
import { AppConfigService } from '../config/app-config.service';
import { CreditGateService } from './credit-gate.service';

export type CheckoutPlanTier = 'solo' | 'firm';

/**
 * Creates Stripe Checkout Sessions for the two subscription tiers priced in
 * cellix-pricing-v3.html (Solo ₹1,299/mo, Firm ₹5,999/mo). Deliberately
 * narrow — CREDIT_SYSTEM.md §7 scopes the rest of the Stripe integration
 * (top-up checkout, billing portal, full webhook-driven fulfillment) as
 * follow-up work; this is the one call the checkout page needs to redirect
 * a user to Stripe's hosted payment page.
 */
@Injectable()
export class StripeCheckoutService {
  private stripeClient: Stripe | undefined;

  constructor(
    private readonly config: AppConfigService,
    private readonly creditGate: CreditGateService,
  ) {}

  private get stripe(): Stripe {
    if (!this.stripeClient) {
      const secretKey = this.config.stripeSecretKey;
      if (!secretKey) {
        throw new ServiceUnavailableException('STRIPE_NOT_CONFIGURED');
      }
      this.stripeClient = new Stripe(secretKey);
    }
    return this.stripeClient;
  }

  private priceIdFor(planTier: CheckoutPlanTier): string {
    const priceId =
      planTier === 'solo' ? this.config.stripePriceSoloMonthly : this.config.stripePriceFirmMonthly;
    if (!priceId) {
      throw new ServiceUnavailableException('STRIPE_PRICE_NOT_CONFIGURED');
    }
    return priceId;
  }

  /**
   * Creates a subscription Checkout Session for billingEntityId and returns
   * its hosted URL. Ensures a credit_accounts document exists first (a user
   * can reach checkout without ever having made a billable request) so the
   * webhook has somewhere to grant plan credits onto once payment succeeds.
   */
  async createSubscriptionSession(
    billingEntityId: string,
    email: string | undefined,
    planTier: CheckoutPlanTier,
  ): Promise<{ url: string }> {
    if (planTier !== 'solo' && planTier !== 'firm') {
      throw new BadRequestException('INVALID_PLAN_TIER');
    }
    await this.creditGate.ensureAccount(billingEntityId);

    const session = await this.stripe.checkout.sessions.create({
      mode: 'subscription',
      line_items: [{ price: this.priceIdFor(planTier), quantity: 1 }],
      success_url: `${this.config.checkoutSuccessUrl}?session_id={CHECKOUT_SESSION_ID}`,
      cancel_url: this.config.checkoutCancelUrl,
      customer_email: email,
      client_reference_id: billingEntityId,
      metadata: { billingEntityId, planTier },
      subscription_data: { metadata: { billingEntityId, planTier } },
    });

    if (!session.url) {
      throw new ServiceUnavailableException('STRIPE_SESSION_NO_URL');
    }
    return { url: session.url };
  }

  /**
   * Entry point for the marketing site's pre-signup checkout — a visitor who
   * has never authenticated against cellix_backend and so has no real
   * userId. The normalized email itself becomes billingEntityId (still
   * `billingEntityType: 'user'` — CreditGateService.ensureAccount's
   * default). This is a deliberate simplification: when this same person
   * later signs into the product via Google/Microsoft OAuth with the same
   * email, reconciling that real userId onto this pre-existing email-keyed
   * account is a follow-up piece of work, not handled here.
   */
  async createGuestSubscriptionSession(email: string, planTier: CheckoutPlanTier): Promise<{ url: string }> {
    const normalizedEmail = normalizeGuestEmail(email);
    if (!normalizedEmail) {
      throw new BadRequestException('INVALID_EMAIL');
    }
    return this.createSubscriptionSession(normalizedEmail, normalizedEmail, planTier);
  }
}

/** Lowercased, trimmed — so "CA@Example.com" and "ca@example.com" resolve to one account. */
function normalizeGuestEmail(email: string): string | null {
  const trimmed = email.trim().toLowerCase();
  return trimmed.length > 0 ? trimmed : null;
}
