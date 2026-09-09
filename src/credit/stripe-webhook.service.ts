import { BadRequestException, Injectable, Logger, ServiceUnavailableException } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import Stripe from 'stripe';
import { AppConfigService } from '../config/app-config.service';
import { ProcessedStripeEvent, ProcessedStripeEventDocument } from './schemas/subscription.schema';
import { Subscription, SubscriptionDocument } from './schemas/subscription.schema';
import { CreditLedgerService } from './credit-ledger.service';
import { CreditGateService } from './credit-gate.service';
import { CreditAccount, CreditAccountDocument } from './schemas/credit-account.schema';

/**
 * CREDIT_SYSTEM.md CD-8's monthly plan allotments.
 *
 * Solo raised 500 -> 1100 (Sept 9, 2026): target COGS ceiling is $5/user/month
 * (out of an ~$8 total per-user budget, the rest going to infra/support/
 * overhead) — at the ~₹0.35/credit blended COGS cellix-pricing-v3.html
 * assumes (~$0.0042/credit at ~₹83/USD), 500 credits only spent ~$2.11 of
 * that budget (42%), leaving real headroom. 1100 credits spends ~$4.64,
 * inside the ceiling with a buffer for token-price drift. Firm's 3,000
 * (pooled across seats, CD-7) is intentionally NOT changed here — its
 * per-seat economics were scoped as a separate decision, not yet made.
 *
 * cellix-pricing-v3.html's marketing copy (hero number, 78%/89% utilization
 * examples, the "512 of 500 credits" overdraft example) still says 500 and
 * has NOT been rewritten to match — its narrative was built around specific
 * usage-scenario math tied to that number, not a value to blind-swap. This
 * constant is the one functional source of truth for what a Solo subscriber
 * actually receives; the page copy is flagged as a known follow-up.
 */
const PLAN_MONTHLY_CREDITS: Record<'solo' | 'firm', number> = {
  solo: 1100,
  firm: 3000,
};

/**
 * Verifies and processes Stripe webhook events. CREDIT_SYSTEM.md §7 still
 * scopes the *full* Stripe integration (top-up checkout, billing portal,
 * dunning/grace-period handling — CREDIT_SYSTEM.md §8 Q3) as follow-up work;
 * this covers the one event the subscribe-checkout flow needs end-to-end —
 * `checkout.session.completed` — plus idempotency for every event type.
 */
@Injectable()
export class StripeWebhookService {
  private readonly logger = new Logger(StripeWebhookService.name);
  private stripeClient: Stripe | undefined;

  constructor(
    private readonly config: AppConfigService,
    @InjectModel(ProcessedStripeEvent.name)
    private readonly processedEventModel: Model<ProcessedStripeEventDocument>,
    @InjectModel(Subscription.name)
    private readonly subscriptionModel: Model<SubscriptionDocument>,
    @InjectModel(CreditAccount.name)
    private readonly creditAccountModel: Model<CreditAccountDocument>,
    private readonly creditGate: CreditGateService,
    private readonly creditLedger: CreditLedgerService,
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

  /**
   * Verifies the raw request body against Stripe's signature header before
   * trusting anything in it — an unverified payload is just JSON anyone
   * could POST, not proof a payment happened.
   */
  verifyAndParseEvent(rawBody: Buffer, signatureHeader: string | undefined): Stripe.Event {
    const webhookSecret = this.config.stripeWebhookSecret;
    if (!webhookSecret) {
      throw new ServiceUnavailableException('STRIPE_WEBHOOK_NOT_CONFIGURED');
    }
    if (!signatureHeader) {
      throw new BadRequestException('MISSING_STRIPE_SIGNATURE');
    }
    try {
      return this.stripe.webhooks.constructEvent(rawBody, signatureHeader, webhookSecret);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'signature verification failed';
      throw new BadRequestException(`INVALID_STRIPE_SIGNATURE: ${message}`);
    }
  }

  /**
   * Idempotent processing: a duplicate delivery of an already-processed
   * event is a no-op, never a double grant. CREDIT_SYSTEM_SCHEMA.md §4.
   */
  async handleVerifiedEvent(event: Stripe.Event): Promise<{ alreadyProcessed: boolean }> {
    const existing = await this.processedEventModel.findOne({ stripeEventId: event.id }).lean();
    if (existing) {
      return { alreadyProcessed: true };
    }

    switch (event.type) {
      case 'checkout.session.completed':
        await this.handleCheckoutSessionCompleted(event.data.object as Stripe.Checkout.Session);
        break;
      default:
        // Every other event type (subscription.updated/deleted, invoice.paid
        // renewals, etc.) is CREDIT_SYSTEM.md §7 follow-up work — logged, not
        // silently dropped, so a gap here is visible rather than invisible.
        this.logger.log(`Stripe event ${event.type} received but not yet handled (${event.id})`);
    }

    await this.processedEventModel.create({ stripeEventId: event.id, processedAt: new Date() });
    return { alreadyProcessed: false };
  }

  /** Backward-compatible entry point for a caller that already has a raw event id (tests, idempotency-only checks). */
  async handleEvent(stripeEventId: string): Promise<{ alreadyProcessed: boolean }> {
    const existing = await this.processedEventModel.findOne({ stripeEventId }).lean();
    if (existing) {
      return { alreadyProcessed: true };
    }
    await this.processedEventModel.create({ stripeEventId, processedAt: new Date() });
    return { alreadyProcessed: false };
  }

  private async handleCheckoutSessionCompleted(session: Stripe.Checkout.Session): Promise<void> {
    const billingEntityId = session.client_reference_id ?? session.metadata?.billingEntityId;
    const planTier = session.metadata?.planTier as 'solo' | 'firm' | undefined;
    const stripeSubscriptionId =
      typeof session.subscription === 'string' ? session.subscription : session.subscription?.id;
    const stripeCustomerId = typeof session.customer === 'string' ? session.customer : session.customer?.id;

    if (!billingEntityId || !planTier || !PLAN_MONTHLY_CREDITS[planTier]) {
      this.logger.error(
        `checkout.session.completed missing billingEntityId/planTier (session ${session.id}) — cannot grant credits`,
      );
      return;
    }

    await this.creditGate.ensureAccount(billingEntityId);
    await this.creditAccountModel.updateOne({ billingEntityId }, { $set: { planTier } });

    if (stripeSubscriptionId && stripeCustomerId) {
      const currentPeriodEnd = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
      await this.subscriptionModel.updateOne(
        { stripeSubscriptionId },
        {
          $set: {
            billingEntityId,
            stripeCustomerId,
            stripeSubscriptionId,
            status: 'active',
            planTier,
            currentPeriodEnd,
            cancelAtPeriodEnd: false,
          },
        },
        { upsert: true },
      );
    }

    await this.creditLedger.grantPlanCredits(billingEntityId, PLAN_MONTHLY_CREDITS[planTier], session.id);
  }
}
