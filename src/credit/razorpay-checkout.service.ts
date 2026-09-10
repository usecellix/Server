import { BadRequestException, Injectable, ServiceUnavailableException } from '@nestjs/common';
import Razorpay from 'razorpay';
import { AppConfigService } from '../config/app-config.service';
import { CreditGateService } from './credit-gate.service';
import { TOPUP_PACKS, TopupPackId } from './topup-packs';

export type CheckoutPlanTier = 'solo' | 'firm' | 'beta';

/**
 * Creates Razorpay Subscriptions for the three subscription tiers priced in
 * cellix-pricing-v3.html (Beta ₹899/mo, Solo ₹1,299/mo, Firm ₹5,999/mo), plus
 * one-time Payment Links for credit top-up packs. Replaces the earlier Stripe
 * integration (TASKS.md #181) — the original pricing doc always specified
 * Razorpay ("Razorpay subscriptions from Day 1"); Stripe was what actually
 * got built first, and this migration brings the code in line with that
 * original intent (TASKS.md, credit-system-v2 session).
 *
 * Razorpay Plans (the `plan_id` each subscription references) are NOT
 * created here — unlike a dynamic Stripe Price, a Razorpay Plan is a
 * one-time setup step (via the dashboard or a setup script) that must exist
 * before this service can create a subscription against it. See
 * RAZORPAY_PLAN_ID_SOLO/FIRM/BETA in AppConfigService.
 */
@Injectable()
export class RazorpayCheckoutService {
  private razorpayClient: Razorpay | undefined;

  constructor(
    private readonly config: AppConfigService,
    private readonly creditGate: CreditGateService,
  ) {}

  private get razorpay(): Razorpay {
    if (!this.razorpayClient) {
      const keyId = this.config.razorpayKeyId;
      const keySecret = this.config.razorpayKeySecret;
      if (!keyId || !keySecret) {
        throw new ServiceUnavailableException('RAZORPAY_NOT_CONFIGURED');
      }
      this.razorpayClient = new Razorpay({ key_id: keyId, key_secret: keySecret });
    }
    return this.razorpayClient;
  }

  private planIdFor(planTier: CheckoutPlanTier): string {
    const planId =
      planTier === 'solo'
        ? this.config.razorpayPlanIdSolo
        : planTier === 'firm'
          ? this.config.razorpayPlanIdFirm
          : this.config.razorpayPlanIdBeta;
    if (!planId) {
      throw new ServiceUnavailableException('RAZORPAY_PLAN_NOT_CONFIGURED');
    }
    return planId;
  }

  /**
   * Creates a Razorpay Subscription for billingEntityId and returns its
   * hosted `short_url` — the single link that walks the customer through
   * mandate authorization (UPI Autopay / card) AND the first charge in one
   * flow, no separate authorization step needed. Ensures a credit_accounts
   * document exists first (a user can reach checkout without ever having
   * made a billable request) so the webhook has somewhere to grant plan
   * credits onto once the subscription activates.
   *
   * `total_count: 120` (10 years of monthly cycles) is a practical stand-in
   * for "renews until cancelled" — Razorpay Subscriptions require a finite
   * cycle count, unlike Stripe's open-ended subscription objects.
   */
  async createSubscriptionSession(
    billingEntityId: string,
    email: string | undefined,
    planTier: CheckoutPlanTier,
  ): Promise<{ url: string }> {
    if (planTier !== 'solo' && planTier !== 'firm' && planTier !== 'beta') {
      throw new BadRequestException('INVALID_PLAN_TIER');
    }
    await this.creditGate.ensureAccount(billingEntityId);

    const subscription = await this.razorpay.subscriptions.create({
      plan_id: this.planIdFor(planTier),
      total_count: 120,
      customer_notify: 1,
      notes: { billingEntityId, planTier, email: email ?? '' },
    });

    if (!subscription.short_url) {
      throw new ServiceUnavailableException('RAZORPAY_SUBSCRIPTION_NO_URL');
    }
    return { url: subscription.short_url };
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

  /**
   * Creates a one-time Razorpay Payment Link for a top-up pack and returns
   * its hosted `short_url`. Authed-only (no guest equivalent) — a top-up
   * implies an existing account/balance to add to, unlike guest subscription
   * signup. Uses inline pack definitions (TOPUP_PACKS) rather than
   * pre-created Razorpay Plans, since a one-time payment doesn't need one.
   */
  async createTopupSession(
    billingEntityId: string,
    email: string | undefined,
    packId: TopupPackId,
  ): Promise<{ url: string }> {
    const pack = TOPUP_PACKS[packId];
    if (!pack) {
      throw new BadRequestException('INVALID_TOPUP_PACK');
    }
    await this.creditGate.ensureAccount(billingEntityId);

    const paymentLink = await this.razorpay.paymentLink.create({
      amount: pack.priceInr * 100,
      currency: 'INR',
      description: `${pack.credits} Cellix credits`,
      customer: email ? { email } : {},
      notify: { email: Boolean(email), sms: false },
      notes: { billingEntityId, packId, credits: pack.credits },
      callback_url: this.config.checkoutSuccessUrl,
      callback_method: 'get',
    });

    if (!paymentLink.short_url) {
      throw new ServiceUnavailableException('RAZORPAY_PAYMENT_LINK_NO_URL');
    }
    return { url: paymentLink.short_url };
  }
}

/** Lowercased, trimmed — so "CA@Example.com" and "ca@example.com" resolve to one account. */
function normalizeGuestEmail(email: string): string | null {
  const trimmed = email.trim().toLowerCase();
  return trimmed.length > 0 ? trimmed : null;
}
