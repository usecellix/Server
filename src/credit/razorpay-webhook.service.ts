import { BadRequestException, Injectable, Logger, ServiceUnavailableException } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import Razorpay from 'razorpay';
import type { Invoices } from 'razorpay/dist/types/invoices';
import type { Subscriptions } from 'razorpay/dist/types/subscriptions';
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
 * `payment.authorized`/`payment.captured` for a Subscription's charge.
 * Razorpay does not embed `notes` or a `subscription_id` directly on the
 * payment entity here — only `invoice_id`, which has to be resolved via
 * `invoices.fetch` to reach the subscription (and its notes) the payment
 * belongs to. See handlePaymentCredited's docblock for why this path exists
 * at all: `subscription.activated` is not guaranteed to fire for every
 * subscription authorization (observed directly — a real UPI-intent-flow
 * Beta subscription in test mode delivered payment.authorized/
 * payment.captured but never subscription.activated, even on retry).
 */
interface RazorpayPaymentWebhookPayload {
  event: string;
  payload: {
    payment?: {
      entity: {
        id: string;
        status: string;
        order_id?: string | null;
        invoice_id?: string | null;
        notes?: Record<string, string | number> | unknown[];
      };
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
  private razorpayClient: Razorpay | undefined;

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

  /** Only needed by handlePaymentCredited's invoice/subscription lookup — every other handler works off the webhook payload alone. */
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
  ): RazorpaySubscriptionWebhookPayload | RazorpayPaymentLinkWebhookPayload | RazorpayPaymentWebhookPayload {
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
      return JSON.parse(bodyText) as
        | RazorpaySubscriptionWebhookPayload
        | RazorpayPaymentLinkWebhookPayload
        | RazorpayPaymentWebhookPayload;
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
    payload: RazorpaySubscriptionWebhookPayload | RazorpayPaymentLinkWebhookPayload | RazorpayPaymentWebhookPayload,
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
      case 'payment.captured':
        // Fallback path — see handlePaymentCredited's docblock. Only
        // 'captured' (money actually settled), not 'authorized' (can still
        // fail to capture), grants credits — same real-funds bar
        // subscription.activated/charged implicitly clear by only firing
        // once Razorpay itself considers the charge successful.
        await this.handlePaymentCredited(payload as RazorpayPaymentWebhookPayload, eventId);
        break;
      default:
        // Every other event type (subscription.updated/pending/paused/
        // resumed, payment_link.partially_paid/cancelled/expired,
        // payment.authorized/failed) is logged, not silently dropped, so a
        // gap here is visible rather than invisible.
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

  private resolveEventId(
    payload: RazorpaySubscriptionWebhookPayload | RazorpayPaymentLinkWebhookPayload | RazorpayPaymentWebhookPayload,
  ): string {
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
    const paymentPayload = payload as RazorpayPaymentWebhookPayload;
    if (paymentPayload.payload.payment) {
      // The bare payment id alone is unique per real occurrence (Razorpay
      // never reuses a payment id across a retry of the SAME payment) —
      // unlike subscription/payment-link ids, which recur across their
      // lifecycle events, this needs no compound key.
      return `${payload.event}:${paymentPayload.payload.payment.entity.id}`;
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
    await this.grantSubscriptionCredits(entity, eventId, payload.event);
  }

  /**
   * Shared grant core for a subscription's successful charge, regardless of
   * which webhook event surfaced it — subscription.activated/charged
   * (normal path) or payment.captured (fallback path, see
   * handlePaymentCredited). Keeping this in one place means the two paths
   * can never grant a different amount or skip a step relative to each
   * other.
   */
  private async grantSubscriptionCredits(
    entity: {
      id: string;
      status: string;
      customer_id: string | null;
      current_start?: number | null;
      current_end?: number | null;
      notes?: Record<string, string | number>;
    },
    eventId: string,
    eventLabel: string,
  ): Promise<void> {
    const billingEntityId = entity.notes?.billingEntityId as string | undefined;
    const planTier = entity.notes?.planTier as PlanTier | undefined;

    if (!billingEntityId || !planTier || !PLAN_MONTHLY_CREDITS[planTier]) {
      this.logger.error(
        `${eventLabel} missing billingEntityId/planTier in notes (subscription ${entity.id}) — cannot grant credits`,
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

  /**
   * Fallback for a Subscription's charge that Razorpay reports ONLY via
   * payment.captured, without ever sending subscription.activated —
   * reproduced directly against a real test-mode Beta subscription paid by
   * UPI intent flow (payment.authorized + payment.captured both delivered;
   * subscription.activated never arrived, even after 40+ minutes, so this is
   * not a delivery-order/retry timing issue). Without this fallback, a real,
   * successfully captured payment silently grants nothing — exactly what
   * happened before this was added.
   *
   * The payment entity itself carries no notes/subscription_id (Razorpay
   * only put `notes: []` — empty — on the payment here), so the only way
   * back to billingEntityId/planTier is: payment.invoice_id ->
   * invoices.fetch -> invoice.subscription_id -> subscriptions.fetch ->
   * subscription.notes. Two extra Razorpay API calls, only paid on this
   * fallback path — the normal subscription.activated/charged path never
   * needs them since the subscription entity is already inline in that
   * webhook's payload.
   *
   * A payment with no invoice_id (e.g. a one-off Order/Payment unrelated to
   * any subscription) is not this service's concern and is left to the
   * `default` logger.log in handleVerifiedEvent — returning early here
   * rather than erroring, since "not every payment is a subscription
   * payment" is expected, not a fault.
   */
  private async handlePaymentCredited(payload: RazorpayPaymentWebhookPayload, eventId: string): Promise<void> {
    const entity = payload.payload.payment?.entity;
    if (!entity) {
      this.logger.error(`${payload.event} missing payment entity — cannot resolve subscription`);
      return;
    }
    if (!entity.invoice_id) {
      // Ordinary one-off payment, not a subscription charge — nothing to grant.
      return;
    }

    let subscriptionId: string | undefined;
    try {
      const invoice = await this.razorpay.invoices.fetch(entity.invoice_id);
      // Razorpay's actual API response includes subscription_id on an
      // invoice entity (razorpay.com/docs/api/payments/invoices/#fetch-an-
      // invoice-by-id) — the SDK's RazorpayInvoice type just never declared
      // it (only RazorpayInvoiceQuery, a LIST filter, has it), so this reads
      // past a real types-package gap rather than an actual missing field.
      subscriptionId = (invoice as Invoices.RazorpayInvoice & { subscription_id?: string }).subscription_id;
    } catch (error) {
      this.logger.error(
        `${payload.event} failed to fetch invoice ${entity.invoice_id} for payment ${entity.id}: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      return;
    }
    if (!subscriptionId) {
      // Invoice exists but isn't tied to a subscription — same "not our concern" case as no invoice_id at all.
      return;
    }

    let subscription: Subscriptions.RazorpaySubscription;
    try {
      subscription = await this.razorpay.subscriptions.fetch(subscriptionId);
    } catch (error) {
      this.logger.error(
        `${payload.event} failed to fetch subscription ${subscriptionId} for payment ${entity.id}: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      return;
    }

    await this.grantSubscriptionCredits(
      {
        id: subscription.id,
        status: subscription.status,
        customer_id: subscription.customer_id,
        current_start: subscription.current_start,
        current_end: subscription.current_end,
        notes: subscription.notes as Record<string, string | number> | undefined,
      },
      eventId,
      payload.event,
    );
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
