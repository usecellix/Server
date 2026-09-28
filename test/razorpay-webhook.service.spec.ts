import { BadRequestException, ServiceUnavailableException } from '@nestjs/common';
import Razorpay from 'razorpay';
import { RazorpayWebhookService } from '../src/credit/razorpay-webhook.service';

function buildService(options: {
  existingEvent?: { paymentEventId: string } | null;
  razorpayWebhookSecret?: string;
  invoicesFetchImpl?: (...args: unknown[]) => unknown;
  subscriptionsFetchImpl?: (...args: unknown[]) => unknown;
} = {}) {
  const insertedEvents: Record<string, unknown>[] = [];
  const processedEventModel = {
    findOne: jest.fn(() => ({ lean: () => Promise.resolve(options.existingEvent ?? null) })),
    create: jest.fn((doc: Record<string, unknown>) => {
      insertedEvents.push(doc);
      return Promise.resolve(doc);
    }),
  };

  const subscriptionUpdates: Record<string, unknown>[] = [];
  const subscriptionModel = {
    updateOne: jest.fn((filter: unknown, update: unknown) => {
      subscriptionUpdates.push({ filter, update });
      return Promise.resolve();
    }),
  };

  const accountUpdates: Record<string, unknown>[] = [];
  const creditAccountModel = {
    updateOne: jest.fn((filter: unknown, update: unknown) => {
      accountUpdates.push({ filter, update });
      return Promise.resolve();
    }),
  };

  const creditGate = { ensureAccount: jest.fn().mockResolvedValue({}) };
  const grantPlanCredits = jest.fn().mockResolvedValue(undefined);
  const addPurchasedCredits = jest.fn().mockResolvedValue(undefined);
  const creditLedger = { grantPlanCredits, addPurchasedCredits };

  const config = {
    razorpayWebhookSecret: options.razorpayWebhookSecret ?? 'whsec_x',
    razorpayKeyId: 'rzp_test_key',
    razorpayKeySecret: 'rzp_test_secret',
  };

  const service = new RazorpayWebhookService(
    config as never,
    processedEventModel as never,
    subscriptionModel as never,
    creditAccountModel as never,
    creditGate as never,
    creditLedger as never,
  );

  // Same pattern as razorpay-checkout.service.spec.ts: bypass the lazy
  // `razorpay` getter by injecting the client directly, since handlePaymentCredited
  // (the payment.captured fallback) is the only path here that calls the real SDK.
  const invoicesFetch = jest.fn(
    options.invoicesFetchImpl ?? (() => Promise.resolve({ subscription_id: 'sub_from_invoice' })),
  );
  const subscriptionsFetch = jest.fn(
    options.subscriptionsFetchImpl ??
      (() =>
        Promise.resolve({
          id: 'sub_from_invoice',
          status: 'active',
          customer_id: 'cust_1',
          current_start: 1700000000,
          current_end: 1702592000,
          notes: { billingEntityId: 'user-1', planTier: 'beta' },
        })),
  );
  (
    service as unknown as {
      razorpayClient?: { invoices: { fetch: unknown }; subscriptions: { fetch: unknown } };
    }
  ).razorpayClient = {
    invoices: { fetch: invoicesFetch },
    subscriptions: { fetch: subscriptionsFetch },
  } as never;

  return {
    service,
    insertedEvents,
    subscriptionUpdates,
    accountUpdates,
    creditGate,
    grantPlanCredits,
    addPurchasedCredits,
    invoicesFetch,
    subscriptionsFetch,
  };
}

describe('RazorpayWebhookService.verifyAndParseEvent', () => {
  afterEach(() => jest.restoreAllMocks());

  it('throws when RAZORPAY_WEBHOOK_SECRET is not configured', () => {
    const { service } = buildService({ razorpayWebhookSecret: '' });
    expect(() => service.verifyAndParseEvent(Buffer.from('{}'), 'sig')).toThrow(ServiceUnavailableException);
  });

  it('throws when the x-razorpay-signature header is missing', () => {
    const { service } = buildService();
    expect(() => service.verifyAndParseEvent(Buffer.from('{}'), undefined)).toThrow(BadRequestException);
  });

  it('throws when signature verification returns false', () => {
    jest.spyOn(Razorpay, 'validateWebhookSignature').mockReturnValue(false);
    const { service } = buildService();
    expect(() => service.verifyAndParseEvent(Buffer.from('{}'), 'sig')).toThrow(BadRequestException);
  });

  it('throws when signature verification itself throws', () => {
    jest.spyOn(Razorpay, 'validateWebhookSignature').mockImplementation(() => {
      throw new Error('bad signature');
    });
    const { service } = buildService();
    expect(() => service.verifyAndParseEvent(Buffer.from('{}'), 'sig')).toThrow(BadRequestException);
  });

  it('parses and returns the body when verification succeeds', () => {
    jest.spyOn(Razorpay, 'validateWebhookSignature').mockReturnValue(true);
    const { service } = buildService();
    const body = Buffer.from(JSON.stringify({ event: 'subscription.activated', payload: {} }));
    expect(service.verifyAndParseEvent(body, 'sig')).toEqual({
      event: 'subscription.activated',
      payload: {},
    });
  });

  it('throws on a verified but malformed (non-JSON) body', () => {
    jest.spyOn(Razorpay, 'validateWebhookSignature').mockReturnValue(true);
    const { service } = buildService();
    expect(() => service.verifyAndParseEvent(Buffer.from('not json'), 'sig')).toThrow(BadRequestException);
  });
});

describe('RazorpayWebhookService.handleVerifiedEvent — subscription.activated/subscription.charged', () => {
  it('is a no-op for a duplicate delivery of an already-processed event', async () => {
    const { service, insertedEvents, grantPlanCredits } = buildService({
      existingEvent: { paymentEventId: 'subscription.activated:sub_1:pay_1' },
    });

    const result = await service.handleVerifiedEvent({
      event: 'subscription.activated',
      payload: {
        subscription: { entity: { id: 'sub_1', plan_id: 'plan_solo', status: 'active', customer_id: 'cust_1' } },
        payment: { entity: { id: 'pay_1' } },
      },
    });

    expect(result).toEqual({ alreadyProcessed: true });
    expect(insertedEvents).toEqual([]);
    expect(grantPlanCredits).not.toHaveBeenCalled();
  });

  it('grants Solo plan credits (3000) on subscription.activated', async () => {
    const { service, insertedEvents, grantPlanCredits, creditGate, accountUpdates } = buildService();

    const result = await service.handleVerifiedEvent({
      event: 'subscription.activated',
      payload: {
        subscription: {
          entity: {
            id: 'sub_1',
            plan_id: 'plan_solo',
            status: 'active',
            customer_id: 'cust_1',
            current_start: 1700000000,
            current_end: 1702592000,
            notes: { billingEntityId: 'user-1', planTier: 'solo' },
          },
        },
        payment: { entity: { id: 'pay_1' } },
      },
    });

    expect(result).toEqual({ alreadyProcessed: false });
    expect(creditGate.ensureAccount).toHaveBeenCalledWith('user-1');
    expect(accountUpdates[0]).toEqual(expect.objectContaining({ filter: { billingEntityId: 'user-1' } }));
    expect(grantPlanCredits).toHaveBeenCalledWith('user-1', 3000, 'subscription.activated:sub_1:pay_1');
    expect(insertedEvents).toEqual([
      expect.objectContaining({ paymentEventId: 'subscription.activated:sub_1:pay_1' }),
    ]);
  });

  it('grants Firm plan credits (3000) on subscription.activated', async () => {
    const { service, grantPlanCredits } = buildService();

    await service.handleVerifiedEvent({
      event: 'subscription.activated',
      payload: {
        subscription: {
          entity: { id: 'sub_2', plan_id: 'plan_firm', status: 'active', customer_id: 'cust_2', notes: { billingEntityId: 'org-1', planTier: 'firm' } },
        },
        payment: { entity: { id: 'pay_2' } },
      },
    });

    expect(grantPlanCredits).toHaveBeenCalledWith('org-1', 3000, expect.any(String));
  });

  it('grants Beta plan credits (500) on subscription.activated', async () => {
    const { service, grantPlanCredits } = buildService();

    await service.handleVerifiedEvent({
      event: 'subscription.activated',
      payload: {
        subscription: {
          entity: { id: 'sub_3', plan_id: 'plan_beta', status: 'active', customer_id: 'cust_3', notes: { billingEntityId: 'user-3', planTier: 'beta' } },
        },
        payment: { entity: { id: 'pay_3' } },
      },
    });

    expect(grantPlanCredits).toHaveBeenCalledWith('user-3', 500, expect.any(String));
  });

  it('ALSO grants plan credits on subscription.charged (renewal) — the gap the Stripe integration never had', async () => {
    // TASKS.md's credit-system-v2 entry: the earlier Stripe integration only
    // ever granted credits ONCE, at initial checkout.session.completed —
    // there was no renewal-grant call-in at all. subscription.charged fires
    // on every subsequent billing cycle and must grant too, or a Solo/Firm/
    // Beta subscriber's credits would never actually replenish.
    const { service, grantPlanCredits } = buildService();

    await service.handleVerifiedEvent({
      event: 'subscription.charged',
      payload: {
        subscription: {
          entity: { id: 'sub_1', plan_id: 'plan_solo', status: 'active', customer_id: 'cust_1', notes: { billingEntityId: 'user-1', planTier: 'solo' } },
        },
        payment: { entity: { id: 'pay_renewal_1' } },
      },
    });

    expect(grantPlanCredits).toHaveBeenCalledWith('user-1', 3000, 'subscription.charged:sub_1:pay_renewal_1');
  });

  it('uses a DIFFERENT idempotency key for each renewal of the SAME subscription (distinct payment ids)', async () => {
    const { service, grantPlanCredits } = buildService();
    const basePayload = (paymentId: string) => ({
      event: 'subscription.charged' as const,
      payload: {
        subscription: {
          entity: { id: 'sub_1', plan_id: 'plan_solo', status: 'active', customer_id: 'cust_1', notes: { billingEntityId: 'user-1', planTier: 'solo' } },
        },
        payment: { entity: { id: paymentId } },
      },
    });

    await service.handleVerifiedEvent(basePayload('pay_renewal_1'));
    await service.handleVerifiedEvent(basePayload('pay_renewal_2'));

    expect(grantPlanCredits).toHaveBeenCalledTimes(2);
    expect(grantPlanCredits.mock.calls[0][2]).not.toBe(grantPlanCredits.mock.calls[1][2]);
  });

  it('writes a subscription row keyed by razorpaySubscriptionId, using the REAL current_start/current_end from the payload', async () => {
    const { service, subscriptionUpdates } = buildService();

    await service.handleVerifiedEvent({
      event: 'subscription.activated',
      payload: {
        subscription: {
          entity: {
            id: 'sub_4',
            plan_id: 'plan_solo',
            status: 'active',
            customer_id: 'cust_4',
            current_start: 1700000000,
            current_end: 1702592000,
            notes: { billingEntityId: 'user-4', planTier: 'solo' },
          },
        },
        payment: { entity: { id: 'pay_4' } },
      },
    });

    expect(subscriptionUpdates[0].filter).toEqual({ razorpaySubscriptionId: 'sub_4' });
    const update = subscriptionUpdates[0].update as { $set: Record<string, unknown> };
    expect(update.$set.currentPeriodStart).toEqual(new Date(1700000000 * 1000));
    expect(update.$set.currentPeriodEnd).toEqual(new Date(1702592000 * 1000));
  });

  it('does not grant credits when billingEntityId/planTier is missing from notes', async () => {
    const { service, grantPlanCredits, insertedEvents } = buildService();

    const result = await service.handleVerifiedEvent({
      event: 'subscription.activated',
      payload: {
        subscription: { entity: { id: 'sub_bad', plan_id: 'plan_solo', status: 'active', customer_id: null } },
        payment: { entity: { id: 'pay_bad' } },
      },
    });

    expect(grantPlanCredits).not.toHaveBeenCalled();
    // Still marked processed — a malformed payload shouldn't be retried forever.
    expect(insertedEvents).toEqual([
      expect.objectContaining({ paymentEventId: 'subscription.activated:sub_bad:pay_bad' }),
    ]);
    expect(result).toEqual({ alreadyProcessed: false });
  });
});

describe('RazorpayWebhookService.handleVerifiedEvent — subscription status-only transitions', () => {
  it('subscription.cancelled updates status and cancelAtPeriodEnd, grants nothing', async () => {
    const { service, subscriptionUpdates, grantPlanCredits } = buildService();

    await service.handleVerifiedEvent({
      event: 'subscription.cancelled',
      payload: {
        subscription: { entity: { id: 'sub_5', plan_id: 'plan_solo', status: 'cancelled', customer_id: 'cust_5' } },
      },
    });

    expect(grantPlanCredits).not.toHaveBeenCalled();
    expect(subscriptionUpdates[0]).toEqual({
      filter: { razorpaySubscriptionId: 'sub_5' },
      update: { $set: { status: 'cancelled', cancelAtPeriodEnd: true } },
    });
  });

  it('subscription.halted updates status without setting cancelAtPeriodEnd', async () => {
    const { service, subscriptionUpdates } = buildService();

    await service.handleVerifiedEvent({
      event: 'subscription.halted',
      payload: {
        subscription: { entity: { id: 'sub_6', plan_id: 'plan_solo', status: 'halted', customer_id: 'cust_6' } },
      },
    });

    expect(subscriptionUpdates[0].update).toEqual({ $set: { status: 'halted', cancelAtPeriodEnd: false } });
  });
});

describe('RazorpayWebhookService.handleVerifiedEvent — payment_link.paid (top-up)', () => {
  it('grants purchased credits matching the pack, not whatever notes.credits says', async () => {
    const { service, addPurchasedCredits, creditGate } = buildService();

    const result = await service.handleVerifiedEvent({
      event: 'payment_link.paid',
      payload: {
        payment_link: {
          entity: { id: 'plink_1', notes: { billingEntityId: 'user-1', packId: 'medium', credits: 1000 } },
        },
        payment: { entity: { id: 'pay_topup_1' } },
      },
    });

    expect(creditGate.ensureAccount).toHaveBeenCalledWith('user-1');
    expect(addPurchasedCredits).toHaveBeenCalledWith('user-1', 1000, 'payment_link.paid:plink_1:pay_topup_1');
    expect(result).toEqual({ alreadyProcessed: false });
  });

  it('trusts TOPUP_PACKS over a tampered/disagreeing notes.credits value', async () => {
    const { service, addPurchasedCredits } = buildService();

    await service.handleVerifiedEvent({
      event: 'payment_link.paid',
      payload: {
        payment_link: {
          entity: { id: 'plink_2', notes: { billingEntityId: 'user-1', packId: 'small', credits: 999999 } },
        },
        payment: { entity: { id: 'pay_topup_2' } },
      },
    });

    // TOPUP_PACKS.small.credits (300), not the tampered 999999.
    expect(addPurchasedCredits).toHaveBeenCalledWith('user-1', 300, expect.any(String));
  });

  it('does not grant credits when billingEntityId/packId is missing or unknown', async () => {
    const { service, addPurchasedCredits } = buildService();

    await service.handleVerifiedEvent({
      event: 'payment_link.paid',
      payload: {
        payment_link: { entity: { id: 'plink_bad', notes: { packId: 'not-a-real-pack' } } },
        payment: { entity: { id: 'pay_bad' } },
      },
    });

    expect(addPurchasedCredits).not.toHaveBeenCalled();
  });
});

describe('RazorpayWebhookService.handleVerifiedEvent — payment.captured (subscription.activated fallback)', () => {
  // Reproduces a real bug: a test-mode Beta subscription paid by UPI intent
  // flow delivered payment.authorized + payment.captured but NEVER
  // subscription.activated, even on retry — so payment.captured has to be
  // able to grant credits on its own by resolving back to the subscription
  // via the payment's invoice_id.
  it('grants credits by resolving invoice_id -> subscription_id -> subscription.notes', async () => {
    const { service, grantPlanCredits, creditGate, invoicesFetch, subscriptionsFetch } = buildService();

    const result = await service.handleVerifiedEvent({
      event: 'payment.captured',
      payload: {
        payment: {
          entity: {
            id: 'pay_1',
            status: 'captured',
            order_id: 'order_1',
            invoice_id: 'inv_1',
            notes: [], // Razorpay sends an empty array here, not an object with billingEntityId
          },
        },
      },
    });

    expect(invoicesFetch).toHaveBeenCalledWith('inv_1');
    expect(subscriptionsFetch).toHaveBeenCalledWith('sub_from_invoice');
    expect(creditGate.ensureAccount).toHaveBeenCalledWith('user-1');
    expect(grantPlanCredits).toHaveBeenCalledWith('user-1', 500, 'payment.captured:pay_1');
    expect(result).toEqual({ alreadyProcessed: false });
  });

  it('is a no-op for a duplicate delivery of the same payment id', async () => {
    const { service, grantPlanCredits } = buildService({
      existingEvent: { paymentEventId: 'payment.captured:pay_1' },
    });

    const result = await service.handleVerifiedEvent({
      event: 'payment.captured',
      payload: { payment: { entity: { id: 'pay_1', status: 'captured', invoice_id: 'inv_1' } } },
    });

    expect(result).toEqual({ alreadyProcessed: true });
    expect(grantPlanCredits).not.toHaveBeenCalled();
  });

  it('does nothing for a payment with no invoice_id (an ordinary one-off payment, not a subscription charge)', async () => {
    const { service, grantPlanCredits, invoicesFetch } = buildService();

    const result = await service.handleVerifiedEvent({
      event: 'payment.captured',
      payload: { payment: { entity: { id: 'pay_2', status: 'captured' } } },
    });

    expect(invoicesFetch).not.toHaveBeenCalled();
    expect(grantPlanCredits).not.toHaveBeenCalled();
    expect(result).toEqual({ alreadyProcessed: false });
  });

  it('does nothing when the invoice exists but has no subscription_id', async () => {
    const { service, grantPlanCredits, subscriptionsFetch } = buildService({
      invoicesFetchImpl: () => Promise.resolve({ subscription_id: undefined }),
    });

    await service.handleVerifiedEvent({
      event: 'payment.captured',
      payload: { payment: { entity: { id: 'pay_3', status: 'captured', invoice_id: 'inv_3' } } },
    });

    expect(subscriptionsFetch).not.toHaveBeenCalled();
    expect(grantPlanCredits).not.toHaveBeenCalled();
  });

  it('does not throw when the invoice lookup itself fails — logs and returns instead', async () => {
    const { service, grantPlanCredits } = buildService({
      invoicesFetchImpl: () => Promise.reject(new Error('Razorpay API down')),
    });

    const result = await service.handleVerifiedEvent({
      event: 'payment.captured',
      payload: { payment: { entity: { id: 'pay_4', status: 'captured', invoice_id: 'inv_4' } } },
    });

    expect(grantPlanCredits).not.toHaveBeenCalled();
    // Still marked processed — a failed API call on redelivery is Razorpay's
    // retry mechanism's job, not something this idempotency layer re-litigates.
    expect(result).toEqual({ alreadyProcessed: false });
  });

  it('does not throw when the subscription lookup itself fails — logs and returns instead', async () => {
    const { service, grantPlanCredits } = buildService({
      subscriptionsFetchImpl: () => Promise.reject(new Error('Razorpay API down')),
    });

    const result = await service.handleVerifiedEvent({
      event: 'payment.captured',
      payload: { payment: { entity: { id: 'pay_5', status: 'captured', invoice_id: 'inv_5' } } },
    });

    expect(grantPlanCredits).not.toHaveBeenCalled();
    expect(result).toEqual({ alreadyProcessed: false });
  });

  it('does not grant credits when the resolved subscription has no billingEntityId/planTier in notes', async () => {
    const { service, grantPlanCredits } = buildService({
      subscriptionsFetchImpl: () =>
        Promise.resolve({ id: 'sub_from_invoice', status: 'active', customer_id: null, notes: {} }),
    });

    await service.handleVerifiedEvent({
      event: 'payment.captured',
      payload: { payment: { entity: { id: 'pay_6', status: 'captured', invoice_id: 'inv_6' } } },
    });

    expect(grantPlanCredits).not.toHaveBeenCalled();
  });

  it('payment.authorized is NOT granted — only payment.captured (real settled funds)', async () => {
    const { service, grantPlanCredits, invoicesFetch } = buildService();

    const result = await service.handleVerifiedEvent({
      event: 'payment.authorized',
      payload: { payment: { entity: { id: 'pay_7', status: 'authorized', invoice_id: 'inv_7' } } },
    });

    expect(invoicesFetch).not.toHaveBeenCalled();
    expect(grantPlanCredits).not.toHaveBeenCalled();
    expect(result).toEqual({ alreadyProcessed: false });
  });
});

describe('RazorpayWebhookService.handleVerifiedEvent — unhandled event types', () => {
  it('marks an unhandled event type as processed without granting credits', async () => {
    const { service, grantPlanCredits, addPurchasedCredits, insertedEvents } = buildService();

    await service.handleVerifiedEvent({
      event: 'subscription.paused',
      payload: {
        subscription: { entity: { id: 'sub_7', plan_id: 'plan_solo', status: 'paused', customer_id: 'cust_7' } },
      },
    });

    expect(grantPlanCredits).not.toHaveBeenCalled();
    expect(addPurchasedCredits).not.toHaveBeenCalled();
    expect(insertedEvents).toEqual([expect.objectContaining({ paymentEventId: expect.stringContaining('subscription.paused') })]);
  });
});

describe('RazorpayWebhookService.handleEvent (idempotency-only entry point)', () => {
  it('reports a duplicate delivery as already processed without re-recording it', async () => {
    const { service, insertedEvents } = buildService({ existingEvent: { paymentEventId: 'evt_1' } });
    const result = await service.handleEvent('evt_1');
    expect(result).toEqual({ alreadyProcessed: true });
    expect(insertedEvents).toEqual([]);
  });

  it('records a first-seen event id', async () => {
    const { service, insertedEvents } = buildService();
    const result = await service.handleEvent('evt_new');
    expect(result).toEqual({ alreadyProcessed: false });
    expect(insertedEvents).toEqual([expect.objectContaining({ paymentEventId: 'evt_new' })]);
  });
});
