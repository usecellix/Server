import { BadRequestException, ServiceUnavailableException } from '@nestjs/common';
import { StripeWebhookService } from '../src/credit/stripe-webhook.service';

function buildService(options: {
  existingEvent?: { stripeEventId: string } | null;
  stripeSecretKey?: string;
  stripeWebhookSecret?: string;
  constructEventImpl?: (...args: unknown[]) => unknown;
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
  const creditLedger = { grantPlanCredits };

  const config = {
    stripeSecretKey: options.stripeSecretKey ?? 'sk_test_x',
    stripeWebhookSecret: options.stripeWebhookSecret ?? 'whsec_x',
  };

  const service = new StripeWebhookService(
    config as never,
    processedEventModel as never,
    subscriptionModel as never,
    creditAccountModel as never,
    creditGate as never,
    creditLedger as never,
  );

  if (options.constructEventImpl) {
    // Reach into the lazily-constructed Stripe client to stub webhooks.constructEvent
    // without hitting the network — same approach as stubbing any lazy singleton.
    (service as unknown as { stripeClient?: { webhooks: { constructEvent: unknown } } }).stripeClient = {
      webhooks: { constructEvent: options.constructEventImpl },
    } as never;
  }

  return { service, insertedEvents, subscriptionUpdates, accountUpdates, creditGate, grantPlanCredits };
}

describe('StripeWebhookService.verifyAndParseEvent', () => {
  it('throws when STRIPE_WEBHOOK_SECRET is not configured', () => {
    const { service } = buildService({ stripeWebhookSecret: '' });
    expect(() => service.verifyAndParseEvent(Buffer.from('{}'), 'sig')).toThrow(ServiceUnavailableException);
  });

  it('throws when the stripe-signature header is missing', () => {
    const { service } = buildService();
    expect(() => service.verifyAndParseEvent(Buffer.from('{}'), undefined)).toThrow(BadRequestException);
  });

  it('throws when signature verification fails', () => {
    const { service } = buildService({
      constructEventImpl: () => {
        throw new Error('bad signature');
      },
    });
    expect(() => service.verifyAndParseEvent(Buffer.from('{}'), 'sig')).toThrow(BadRequestException);
  });

  it('returns the parsed event when verification succeeds', () => {
    const fakeEvent = { id: 'evt_1', type: 'checkout.session.completed' };
    const { service } = buildService({ constructEventImpl: () => fakeEvent });
    expect(service.verifyAndParseEvent(Buffer.from('{}'), 'sig')).toBe(fakeEvent);
  });
});

describe('StripeWebhookService.handleVerifiedEvent', () => {
  it('is a no-op for a duplicate delivery of an already-processed event', async () => {
    const { service, insertedEvents, grantPlanCredits } = buildService({
      existingEvent: { stripeEventId: 'evt_1' },
    });

    const result = await service.handleVerifiedEvent({
      id: 'evt_1',
      type: 'checkout.session.completed',
      data: { object: {} },
    } as never);

    expect(result).toEqual({ alreadyProcessed: true });
    expect(insertedEvents).toEqual([]);
    expect(grantPlanCredits).not.toHaveBeenCalled();
  });

  it('grants Solo plan credits (500) on a completed Solo checkout session', async () => {
    const { service, insertedEvents, grantPlanCredits, creditGate, accountUpdates } = buildService();

    const result = await service.handleVerifiedEvent({
      id: 'evt_solo',
      type: 'checkout.session.completed',
      data: {
        object: {
          id: 'cs_1',
          client_reference_id: 'user-1',
          metadata: { billingEntityId: 'user-1', planTier: 'solo' },
          subscription: 'sub_1',
          customer: 'cus_1',
        },
      },
    } as never);

    expect(result).toEqual({ alreadyProcessed: false });
    expect(creditGate.ensureAccount).toHaveBeenCalledWith('user-1');
    expect(accountUpdates[0]).toEqual(
      expect.objectContaining({ filter: { billingEntityId: 'user-1' } }),
    );
    expect(grantPlanCredits).toHaveBeenCalledWith('user-1', 500, 'cs_1');
    expect(insertedEvents).toEqual([expect.objectContaining({ stripeEventId: 'evt_solo' })]);
  });

  it('grants Firm plan credits (3000) on a completed Firm checkout session', async () => {
    const { service, grantPlanCredits } = buildService();

    await service.handleVerifiedEvent({
      id: 'evt_firm',
      type: 'checkout.session.completed',
      data: {
        object: {
          id: 'cs_2',
          client_reference_id: 'org-1',
          metadata: { billingEntityId: 'org-1', planTier: 'firm' },
          subscription: 'sub_2',
          customer: 'cus_2',
        },
      },
    } as never);

    expect(grantPlanCredits).toHaveBeenCalledWith('org-1', 3000, 'cs_2');
  });

  it('writes a subscription row keyed by stripeSubscriptionId', async () => {
    const { service, subscriptionUpdates } = buildService();

    await service.handleVerifiedEvent({
      id: 'evt_solo2',
      type: 'checkout.session.completed',
      data: {
        object: {
          id: 'cs_3',
          client_reference_id: 'user-2',
          metadata: { billingEntityId: 'user-2', planTier: 'solo' },
          subscription: 'sub_3',
          customer: 'cus_3',
        },
      },
    } as never);

    expect(subscriptionUpdates[0].filter).toEqual({ stripeSubscriptionId: 'sub_3' });
  });

  it('does not grant credits when billingEntityId/planTier is missing from the session', async () => {
    const { service, grantPlanCredits, insertedEvents } = buildService();

    const result = await service.handleVerifiedEvent({
      id: 'evt_bad',
      type: 'checkout.session.completed',
      data: { object: { id: 'cs_4' } },
    } as never);

    expect(grantPlanCredits).not.toHaveBeenCalled();
    // Still marked processed — a malformed session shouldn't be retried forever.
    expect(insertedEvents).toEqual([expect.objectContaining({ stripeEventId: 'evt_bad' })]);
    expect(result).toEqual({ alreadyProcessed: false });
  });

  it('marks an unhandled event type as processed without granting credits', async () => {
    const { service, grantPlanCredits, insertedEvents } = buildService();

    await service.handleVerifiedEvent({
      id: 'evt_other',
      type: 'customer.subscription.updated',
      data: { object: {} },
    } as never);

    expect(grantPlanCredits).not.toHaveBeenCalled();
    expect(insertedEvents).toEqual([expect.objectContaining({ stripeEventId: 'evt_other' })]);
  });
});

describe('StripeWebhookService.handleEvent (idempotency-only entry point)', () => {
  it('reports a duplicate delivery as already processed without re-recording it', async () => {
    const { service, insertedEvents } = buildService({ existingEvent: { stripeEventId: 'evt_1' } });
    const result = await service.handleEvent('evt_1');
    expect(result).toEqual({ alreadyProcessed: true });
    expect(insertedEvents).toEqual([]);
  });

  it('records a first-seen event id', async () => {
    const { service, insertedEvents } = buildService();
    const result = await service.handleEvent('evt_new');
    expect(result).toEqual({ alreadyProcessed: false });
    expect(insertedEvents).toEqual([expect.objectContaining({ stripeEventId: 'evt_new' })]);
  });
});
