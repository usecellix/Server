import { BadRequestException, Injectable, Logger, ServiceUnavailableException } from '@nestjs/common';
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
  private readonly logger = new Logger('RazorpayCheckout');
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
   *
   * Unlike Payment Links, Razorpay's Subscriptions API has no `callback_url`
   * — a `callback_url`/`callback_method` field on `subscriptions.create()`
   * is rejected by Razorpay's API (confirmed against the SDK's own type
   * definitions, which list no such field on the create body), which is why
   * this previously 500/503'd for every signed-in subscriber while guest
   * checkout (which never set that field) succeeded. Without it, the
   * customer lands on Razorpay's own post-payment confirmation screen
   * instead of back on /app — the "always redirect to dashboard" behavior
   * has to be configured from the Razorpay Dashboard itself (Settings →
   * Subscriptions → redirect/return URL) rather than per-API-call; see
   * RAZORPAY_SETUP.md. The webhook remains the source of truth for granting
   * credits regardless of what page the customer ends up on.
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

    try {
      const planId = this.planIdFor(planTier);
      this.logger.log(`Creating subscription: planTier=${planTier}, planId=${planId}, billingEntityId=${billingEntityId}`);

      const subscription = await this.razorpay.subscriptions.create({
        plan_id: planId,
        total_count: 120,
        customer_notify: 1,
        notes: { billingEntityId, planTier, email: email ?? '' },
      });

      if (!subscription.short_url) {
        this.logger.error(`Razorpay subscription created but no short_url: ${JSON.stringify(subscription)}`);
        throw new ServiceUnavailableException('RAZORPAY_SUBSCRIPTION_NO_URL');
      }
      this.logger.log(`Subscription created successfully: url=${subscription.short_url}`);
      return { url: subscription.short_url };
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      const errorBody = (error as any)?.response?.body || (error as any)?.body || {};
      this.logger.error(`Failed to create subscription: ${errorMsg}`, {
        errorMsg,
        errorCode: (error as any)?.code,
        errorBody,
        fullError: error,
      });
      throw new ServiceUnavailableException(
        `Razorpay API error: ${errorMsg}. Check that RAZORPAY_PLAN_ID_${planTier.toUpperCase()} is correct and exists in your test account.`
      );
    }
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

    try {
      this.logger.log(`Creating topup: packId=${packId}, credits=${pack.credits}, priceInr=${pack.priceInr}, billingEntityId=${billingEntityId}`);

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
        this.logger.error(`Razorpay payment link created but no short_url: ${JSON.stringify(paymentLink)}`);
        throw new ServiceUnavailableException('RAZORPAY_PAYMENT_LINK_NO_URL');
      }
      this.logger.log(`Payment link created successfully: url=${paymentLink.short_url}`);
      return { url: paymentLink.short_url };
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      const errorBody = (error as any)?.response?.body || (error as any)?.body || {};
      this.logger.error(`Failed to create topup: ${errorMsg}`, {
        errorMsg,
        errorCode: (error as any)?.code,
        errorBody,
        fullError: error,
      });
      throw new ServiceUnavailableException(
        `Razorpay API error: ${errorMsg}. Check that your Razorpay credentials are correct.`
      );
    }
  }
}

/** Lowercased, trimmed — so "CA@Example.com" and "ca@example.com" resolve to one account. */
function normalizeGuestEmail(email: string): string | null {
  const trimmed = email.trim().toLowerCase();
  return trimmed.length > 0 ? trimmed : null;
}
