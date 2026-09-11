import { BadRequestException, Injectable, Logger, ServiceUnavailableException } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import Razorpay from 'razorpay';
import { AppConfigService } from '../config/app-config.service';
import {
  ProcessedRazorpayEvent,
  ProcessedRazorpayEventDocument,
  Subscription,
  SubscriptionDocument,
} from './schemas/subscription.schema';
import { CreditLedgerService } from './credit-ledger.service';
import { CreditGateService } from './credit-gate.service';
import { CreditAccount, CreditAccountDocument } from './schemas/credit-account.schema';
import { TOPUP_PACKS, TopupPackId } from './topup-packs';

type PlanTier = 'solo' | 'firm' | 'beta';

/**
 * CREDIT_SYSTEM.md CD-8's monthly plan allotments.
 *
 * Re-derived 2026-09-10 against real GLM model costs (~18.7x cheaper per
 * token than the gpt-5 pricing the original ₹0.35/credit blended-COGS
 * assumption was built on) — see TASKS.md's credit-system-v2 entry for the
 * full derivation. Solo raised 1100 -> 3000 after confirming the OLD number
 * couldn't even cover one large (20-subtask) Tier 3 build per month at the
 * new per-subtask Tier 3 pricing (TIER3_AGENTIC_BUILD, credit-cost-catalog.ts).
 * Firm's 3,000 (pooled across seats, CD-7) is intentionally NOT changed here
 * — its per-seat economics and org-billing plumbing are a separate,
 * not-yet-built follow-up. Beta is new: ₹899/mo, 500 credits/mo, deliberately
 * kept fixed (not linked to Solo's number) as a pricing-validation
 * experiment for the first 25 founding members — see cellix-pricing-v3.html
 * Phase 1.
 */
const PLAN_MONTHLY_CREDITS: Record<PlanTier, number> = {
  solo: 3000,
  firm: 3000,
  beta: 500,
};

interface RazorpaySubscriptionWebhookPayload {
  event: string;
  payload: {
    subscription?: {
      entity: {
        id: string;
        plan_id: string;
        status: string;
        customer_id: string | null;
        current_start?: number | null;
        current_end?: number | null;
        notes?: Record<string, string | number>;
      };
    };
    payment?: {
      entity: { id: string };
    };
  };
}

interface RazorpayPaymentLinkWebhookPayload {
  event: string;
  payload: {
    payment_link?: {
      entity: {
        id: string;
        notes?: Record<string, string | number>;
      };
    };
    payment?: {
      entity: { id: string };
    };
  };
}

/**
 * Verifies and processes Razorpay webhook events. Replaces the earlier
 * Stripe integration (TASKS.md #181) — the original pricing doc always
 * specified Razorpay; this migration brings the code in line with that.
 *
 * Closes a real pre-existing gap the Stripe integration never addressed:
 * that integration only ever granted credits ONCE, at initial checkout
 * (`checkout.session.completed`) — there was no renewal-grant logic at all,
 * a known, explicitly-flagged limitation. This service grants credits on
 * BOTH `subscription.activated` (first charge) and `subscription.charged`
 * (every subsequent renewal), so a Solo/Firm/Beta subscriber's credits
 * actually replenish each billing cycle, not just once.
 */
@Injectable()
export class RazorpayWebhookService {
  private readonly logger = new Logger(RazorpayWebhookService.name);

  constructor(
    private readonly config: AppConfigService,
    @InjectModel(ProcessedRazorpayEvent.name)
    private readonly processedEventModel: Model<ProcessedRazorpayEventDocument>,
    @InjectModel(Subscription.name)
    private readonly subscriptionModel: Model<SubscriptionDocument>,
    @InjectModel(CreditAccount.name)
    private readonly creditAccountModel: Model<CreditAccountDocument>,
    private readonly creditGate: CreditGateService,
    private readonly creditLedger: CreditLedgerService,
  ) {}

  /**
   * Verifies the raw request body against Razorpay's HMAC-SHA256 signature
   * header before trusting anything in it, then parses it — an unverified
   * payload is just JSON anyone could POST, not proof a payment happened.
   * Unlike Stripe's `webhooks.constructEvent` (verify + parse in one call),
   * Razorpay's `validateWebhookSignature` only returns a boolean, so parsing
   * is a separate step here.
   */
  verifyAndParseEvent(
    rawBody: Buffer,
    signatureHeader: string | undefined,
  ): RazorpaySubscriptionWebhookPayload | RazorpayPaymentLinkWebhookPayload {
    const webhookSecret = this.config.razorpayWebhookSecret;
    if (!webhookSecret) {
      throw new ServiceUnavailableException('RAZORPAY_WEBHOOK_NOT_CONFIGURED');
    }
    if (!signatureHeader) {
      throw new BadRequestException('MISSING_RAZORPAY_SIGNATURE');
    }
    const bodyText = rawBody.toString('utf8');
    let valid: boolean;
    try {
      valid = Razorpay.validateWebhookSignature(bodyText, signatureHeader, webhookSecret);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'signature verification failed';
      throw new BadRequestException(`INVALID_RAZORPAY_SIGNATURE: ${message}`);
    }
    if (!valid) {
      throw new BadRequestException('INVALID_RAZORPAY_SIGNATURE');
    }
    try {
      return JSON.parse(bodyText) as RazorpaySubscriptionWebhookPayload | RazorpayPaymentLinkWebhookPayload;
    } catch {
      throw new BadRequestException('MALFORMED_RAZORPAY_PAYLOAD');
    }
  }

  /**
   * Idempotent processing: a duplicate delivery of an already-processed
   * event is a no-op, never a double grant. CREDIT_SYSTEM_SCHEMA.md §4.
   *
   * Razorpay's webhook payload has no single top-level `event.id` the way
   * Stripe's does — the idempotency key here is a composite of the event
   * type plus whichever entity id actually changes per real occurrence (the
   * payment id for a subscription charge, since the SAME subscription id
   * recurs every renewal but each charge has its own payment id; the
   * payment-link payment id for a top-up).
   */
  async handleVerifiedEvent(
    payload: RazorpaySubscriptionWebhookPayload | RazorpayPaymentLinkWebhookPayload,
  ): Promise<{ alreadyProcessed: boolean }> {
    const eventId = this.resolveEventId(payload);
    const existing = await this.processedEventModel.findOne({ paymentEventId: eventId }).lean();
    if (existing) {
      return { alreadyProcessed: true };
    }

    switch (payload.event) {
      case 'subscription.activated':
      case 'subscription.charged':
        await this.handleSubscriptionCredited(payload as RazorpaySubscriptionWebhookPayload, eventId);
        break;
      case 'subscription.cancelled':
      case 'subscription.halted':
      case 'subscription.completed':
      case 'subscription.expired':
        await this.handleSubscriptionStatusChange(payload as RazorpaySubscriptionWebhookPayload);
        break;
      case 'payment_link.paid':
        await this.handlePaymentLinkPaid(payload as RazorpayPaymentLinkWebhookPayload, eventId);
        break;
      default:
        // Every other event type (subscription.updated/pending/paused/
        // resumed, payment_link.partially_paid/cancelled/expired) is logged,
        // not silently dropped, so a gap here is visible rather than invisible.
        this.logger.log(`Razorpay event ${payload.event} received but not yet handled (${eventId})`);
    }

    await this.processedEventModel.create({ paymentEventId: eventId, processedAt: new Date() });
    return { alreadyProcessed: false };
  }

  /** Backward-compatible entry point for a caller that already has a resolved event id (tests, idempotency-only checks). */
  async handleEvent(paymentEventId: string): Promise<{ alreadyProcessed: boolean }> {
    const existing = await this.processedEventModel.findOne({ paymentEventId }).lean();
    if (existing) {
      return { alreadyProcessed: true };
    }
    await this.processedEventModel.create({ paymentEventId, processedAt: new Date() });
    return { alreadyProcessed: false };
  }

  private resolveEventId(payload: RazorpaySubscriptionWebhookPayload | RazorpayPaymentLinkWebhookPayload): string {
    const subscriptionPayload = payload as RazorpaySubscriptionWebhookPayload;
    if (subscriptionPayload.payload.subscription) {
      const subscriptionId = subscriptionPayload.payload.subscription.entity.id;
      const paymentId = subscriptionPayload.payload.payment?.entity.id;
      return paymentId
        ? `${payload.event}:${subscriptionId}:${paymentId}`
        : `${payload.event}:${subscriptionId}`;
    }
    const linkPayload = payload as RazorpayPaymentLinkWebhookPayload;
    if (linkPayload.payload.payment_link) {
      const linkId = linkPayload.payload.payment_link.entity.id;
      const paymentId = linkPayload.payload.payment?.entity.id;
      return paymentId ? `${payload.event}:${linkId}:${paymentId}` : `${payload.event}:${linkId}`;
    }
    // Should not happen for a payload Razorpay actually sends — fall back to
    // a coarser key rather than throwing, so an unexpected shape still gets
    // SOME idempotency rather than none.
    return `${payload.event}:${JSON.stringify(payload.payload)}`;
  }

  /** `subscription.activated` (first charge) and `subscription.charged` (every renewal) both grant credits. */
  private async handleSubscriptionCredited(
    payload: RazorpaySubscriptionWebhookPayload,
    eventId: string,
  ): Promise<void> {
    const entity = payload.payload.subscription?.entity;
    if (!entity) {
      this.logger.error(`${payload.event} missing subscription entity — cannot grant credits`);
      return;
    }

    const billingEntityId = entity.notes?.billingEntityId as string | undefined;
    const planTier = entity.notes?.planTier as PlanTier | undefined;

    if (!billingEntityId || !planTier || !PLAN_MONTHLY_CREDITS[planTier]) {
      this.logger.error(
        `${payload.event} missing billingEntityId/planTier in notes (subscription ${entity.id}) — cannot grant credits`,
      );
      return;
    }

    await this.creditGate.ensureAccount(billingEntityId);
    await this.creditAccountModel.updateOne({ billingEntityId }, { $set: { planTier } });

    const currentPeriodStart =
      typeof entity.current_start === 'number' ? new Date(entity.current_start * 1000) : undefined;
    const currentPeriodEnd =
      typeof entity.current_end === 'number' ? new Date(entity.current_end * 1000) : undefined;

    await this.subscriptionModel.updateOne(
      { razorpaySubscriptionId: entity.id },
      {
        $set: {
          billingEntityId,
          razorpayCustomerId: entity.customer_id ?? undefined,
          razorpaySubscriptionId: entity.id,
          status: entity.status,
          planTier,
          ...(currentPeriodStart ? { currentPeriodStart } : {}),
          // Razorpay's current_end can be null pre-activation; fall back to a
          // conservative estimate only when the real value isn't available
          // yet, rather than leaving the field unset (it's required on the
          // schema).
          currentPeriodEnd: currentPeriodEnd ?? new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
          cancelAtPeriodEnd: false,
        },
      },
      { upsert: true },
    );

    await this.creditLedger.grantPlanCredits(billingEntityId, PLAN_MONTHLY_CREDITS[planTier], eventId);
  }

  /** Status-only transitions (cancellation, halting, completion, expiry) — updates the subscription row, grants nothing. */
  private async handleSubscriptionStatusChange(payload: RazorpaySubscriptionWebhookPayload): Promise<void> {
    const entity = payload.payload.subscription?.entity;
    if (!entity) return;

    await this.subscriptionModel.updateOne(
      { razorpaySubscriptionId: entity.id },
      {
        $set: {
          status: entity.status,
          cancelAtPeriodEnd: payload.event === 'subscription.cancelled',
        },
      },
    );
  }

  /** `payment_link.paid` — grants purchased (top-up) credits, never expiring. */
  private async handlePaymentLinkPaid(
    payload: RazorpayPaymentLinkWebhookPayload,
    eventId: string,
  ): Promise<void> {
    const entity = payload.payload.payment_link?.entity;
    if (!entity) {
      this.logger.error('payment_link.paid missing payment_link entity — cannot grant credits');
      return;
    }

    const billingEntityId = entity.notes?.billingEntityId as string | undefined;
    const packId = entity.notes?.packId as TopupPackId | undefined;
    const notedCredits = entity.notes?.credits;

    if (!billingEntityId || !packId || !TOPUP_PACKS[packId]) {
      this.logger.error(
        `payment_link.paid missing billingEntityId/packId in notes (payment link ${entity.id}) — cannot grant credits`,
      );
      return;
    }

    // Prefer the pack's own catalog amount over whatever was echoed back in
    // notes — notes are just round-tripped metadata, not a trusted source of
    // the credit amount, even though this service is also what wrote them.
    const credits = TOPUP_PACKS[packId].credits;
    if (typeof notedCredits === 'number' && notedCredits !== credits) {
      this.logger.warn(
        `payment_link.paid notes.credits (${notedCredits}) disagrees with TOPUP_PACKS[${packId}].credits (${credits}) — using the catalog value`,
      );
    }

    await this.creditGate.ensureAccount(billingEntityId);
    await this.creditLedger.addPurchasedCredits(billingEntityId, credits, eventId);
  }
}
