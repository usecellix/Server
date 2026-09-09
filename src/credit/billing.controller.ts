import { BadRequestException, Body, Controller, Get, Headers, Post, Query, Req, UseGuards } from '@nestjs/common';
import { FastifyRequest } from 'fastify';
import { AuthGuard, AuthUserSession, Session } from '../auth/auth.guard';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { CreditAccountQueryService } from './credit-account-query.service';
import { CreditGateService } from './credit-gate.service';
import { StripeCheckoutService, CheckoutPlanTier } from './stripe-checkout.service';
import { StripeWebhookService } from './stripe-webhook.service';
import { ListLedgerQueryDto } from './dto/list-ledger-query.dto';
import { CreateCheckoutSessionDto } from './dto/create-checkout-session.dto';
import { CreateGuestCheckoutSessionDto } from './dto/create-guest-checkout-session.dto';

/**
 * CREDIT_SYSTEM_SCHEMA.md §5. `/billing/checkout/topup` and `/billing/portal`
 * still need building — CREDIT_SYSTEM.md §7 scopes the rest of the Stripe
 * integration as follow-up work; this covers the one checkout flow the
 * pricing/checkout pages need (subscribe to Solo or Firm).
 *
 * `billingEntityId` is always resolved from the authenticated session, never
 * a request parameter — same reasoning ConversationController's `userId`
 * docblock gives: accepting it as input would let any caller read another
 * user's balance or start a checkout session billed to someone else. Org-
 * pooled billing (CD-7) has no session-to-orgId resolution yet, so every
 * account this controller can address today is `billingEntityType: 'user'`,
 * keyed by `session.user.id`.
 */
@UseGuards(AuthGuard)
@Controller('billing')
export class BillingController {
  constructor(
    private readonly creditAccountQueryService: CreditAccountQueryService,
    private readonly creditGate: CreditGateService,
    private readonly stripeCheckout: StripeCheckoutService,
  ) {}

  @Get('account')
  @SkipEnvelope()
  async getAccount(@Session() session: AuthUserSession) {
    // Provision-on-read (Sept 9, 2026): this used to 404 with
    // NO_CREDIT_ACCOUNT_YET until a user's first BILLABLE action ran the
    // credit gate — meaning the balance chip a brand-new user opens the task
    // pane to check literally could not render, since nothing had spent
    // anything yet. That contradicts CREDIT_SYSTEM.md §5's own "petrol gauge,
    // not a countdown timer" framing: a gauge that's blank until you've
    // already used gas isn't a gauge. `ensureAccount` is the SAME
    // upsert-with-setOnInsert `CreditGateService` already uses before a spend
    // — calling it here just moves "when do we first materialize the
    // Free-tier account" from "first spend" to "first time anyone asks,"
    // which for a real user is signup-adjacent (opening the add-in). Safe
    // under concurrent calls for the same reason the gate's own callers are:
    // `$setOnInsert` makes a race between two simultaneous first-reads a
    // no-op on the loser, not a duplicate account.
    await this.creditGate.ensureAccount(session.user.id);
    const summary = await this.creditAccountQueryService.getAccountSummary(session.user.id);
    return summary;
  }

  @Get('ledger')
  @SkipEnvelope()
  async getLedger(@Session() session: AuthUserSession, @Query() query: ListLedgerQueryDto) {
    return this.creditAccountQueryService.getLedgerPage(session.user.id, query);
  }

  @Post('checkout/subscribe')
  @SkipEnvelope()
  async createSubscribeCheckout(
    @Session() session: AuthUserSession,
    @Body() body: CreateCheckoutSessionDto,
  ) {
    return this.stripeCheckout.createSubscriptionSession(
      session.user.id,
      session.user.email ?? undefined,
      body.planTier as CheckoutPlanTier,
    );
  }
}

/**
 * Unauthenticated checkout entry point for the marketing site
 * (CELLIX-landing-page), where a visitor has never signed into the product
 * and so has no AuthGuard session. Deliberately a separate controller (no
 * `@UseGuards(AuthGuard)`) rather than an `@Session() session?` optional
 * param on BillingController — this keeps the authenticated
 * checkout/subscribe route's security posture untouched and makes the
 * unauthenticated path something a reviewer sees and reasons about
 * explicitly, not a side effect of loosening an existing guard.
 * StripeCheckoutService.createGuestSubscriptionSession's docblock explains
 * why the submitted email becomes billingEntityId.
 */
@Controller('billing/public')
export class PublicBillingController {
  constructor(private readonly stripeCheckout: StripeCheckoutService) {}

  @Post('checkout/subscribe')
  @SkipEnvelope()
  async createGuestSubscribeCheckout(@Body() body: CreateGuestCheckoutSessionDto) {
    return this.stripeCheckout.createGuestSubscriptionSession(
      body.email,
      body.planTier as CheckoutPlanTier,
    );
  }
}

/**
 * Separate controller (no AuthGuard — Stripe calls this, not a logged-in
 * user) so the webhook route's auth posture is never accidentally
 * inherited from BillingController's `@UseGuards(AuthGuard)`.
 */
@Controller('webhooks')
export class StripeWebhookController {
  constructor(private readonly stripeWebhookService: StripeWebhookService) {}

  @Post('stripe')
  @SkipEnvelope()
  async handleStripeWebhook(
    @Req() request: FastifyRequest,
    @Headers('stripe-signature') signature: string | undefined,
  ) {
    const rawBody = (request as { rawBody?: Buffer }).rawBody;
    if (!rawBody) {
      // Nest is started with rawBody: true (main.ts) so JSON requests retain
      // req.rawBody — its absence means that option isn't wired, not a caller error.
      throw new BadRequestException('MISSING_RAW_BODY');
    }
    const event = this.stripeWebhookService.verifyAndParseEvent(rawBody, signature);
    return this.stripeWebhookService.handleVerifiedEvent(event);
  }
}
