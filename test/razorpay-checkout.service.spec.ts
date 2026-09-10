import { BadRequestException, ServiceUnavailableException } from '@nestjs/common';
import { RazorpayCheckoutService } from '../src/credit/razorpay-checkout.service';

function buildService(options: {
  razorpayKeyId?: string;
  razorpayKeySecret?: string;
  razorpayPlanIdSolo?: string;
  razorpayPlanIdFirm?: string;
  razorpayPlanIdBeta?: string;
  createSubscriptionImpl?: (...args: unknown[]) => unknown;
  createPaymentLinkImpl?: (...args: unknown[]) => unknown;
} = {}) {
  const config = {
    razorpayKeyId: options.razorpayKeyId ?? 'rzp_test_key',
    razorpayKeySecret: options.razorpayKeySecret ?? 'rzp_test_secret',
    razorpayPlanIdSolo: options.razorpayPlanIdSolo ?? 'plan_solo',
    razorpayPlanIdFirm: options.razorpayPlanIdFirm ?? 'plan_firm',
    razorpayPlanIdBeta: options.razorpayPlanIdBeta ?? 'plan_beta',
    checkoutSuccessUrl: 'http://localhost:5173/checkout/success',
    checkoutCancelUrl: 'http://localhost:5173/checkout',
  };
  const ensureAccount = jest.fn().mockResolvedValue({});
  const creditGate = { ensureAccount };

  const service = new RazorpayCheckoutService(config as never, creditGate as never);

  const createSubscription = jest.fn(
    options.createSubscriptionImpl ?? (() => Promise.resolve({ short_url: 'https://rzp.io/i/sub_x' })),
  );
  const createPaymentLink = jest.fn(
    options.createPaymentLinkImpl ?? (() => Promise.resolve({ short_url: 'https://rzp.io/i/link_x' })),
  );
  (
    service as unknown as {
      razorpayClient?: {
        subscriptions: { create: unknown };
        paymentLink: { create: unknown };
      };
    }
  ).razorpayClient = {
    subscriptions: { create: createSubscription },
    paymentLink: { create: createPaymentLink },
  } as never;

  return { service, ensureAccount, createSubscription, createPaymentLink };
}

describe('RazorpayCheckoutService.createSubscriptionSession', () => {
  it('rejects a plan tier outside solo/firm/beta', async () => {
    const { service } = buildService();
    await expect(
      service.createSubscriptionSession('user-1', 'ca@example.com', 'enterprise' as never),
    ).rejects.toThrow(BadRequestException);
  });

  it('ensures the account exists before creating the subscription', async () => {
    const { service, ensureAccount } = buildService();
    await service.createSubscriptionSession('user-1', 'ca@example.com', 'solo');
    expect(ensureAccount).toHaveBeenCalledWith('user-1');
  });

  it('creates a subscription with billingEntityId/planTier in notes, not metadata (Razorpay has no metadata field)', async () => {
    const { service, createSubscription } = buildService();
    await service.createSubscriptionSession('user-1', 'ca@example.com', 'solo');

    expect(createSubscription).toHaveBeenCalledWith(
      expect.objectContaining({
        plan_id: 'plan_solo',
        total_count: 120,
        customer_notify: 1,
        notes: { billingEntityId: 'user-1', planTier: 'solo', email: 'ca@example.com' },
      }),
    );
  });

  it('uses the firm plan id for the firm tier', async () => {
    const { service, createSubscription } = buildService();
    await service.createSubscriptionSession('org-1', undefined, 'firm');
    expect(createSubscription).toHaveBeenCalledWith(expect.objectContaining({ plan_id: 'plan_firm' }));
  });

  it('uses the beta plan id for the beta tier', async () => {
    const { service, createSubscription } = buildService();
    await service.createSubscriptionSession('user-1', 'ca@example.com', 'beta');
    expect(createSubscription).toHaveBeenCalledWith(expect.objectContaining({ plan_id: 'plan_beta' }));
  });

  it('returns the subscription short_url', async () => {
    const { service } = buildService();
    const result = await service.createSubscriptionSession('user-1', 'ca@example.com', 'solo');
    expect(result).toEqual({ url: 'https://rzp.io/i/sub_x' });
  });

  it('throws when the subscription has no short_url', async () => {
    const { service } = buildService({ createSubscriptionImpl: () => Promise.resolve({ short_url: undefined }) });
    await expect(service.createSubscriptionSession('user-1', 'ca@example.com', 'solo')).rejects.toThrow(
      ServiceUnavailableException,
    );
  });

  it('throws when the plan id for the plan tier is not configured', async () => {
    const { service } = buildService({ razorpayPlanIdSolo: '' });
    await expect(service.createSubscriptionSession('user-1', 'ca@example.com', 'solo')).rejects.toThrow(
      ServiceUnavailableException,
    );
  });
});

describe('RazorpayCheckoutService.createGuestSubscriptionSession', () => {
  it('normalizes the email (trim + lowercase) into billingEntityId', async () => {
    const { service, createSubscription, ensureAccount } = buildService();
    await service.createGuestSubscriptionSession('  CA@Example.com  ', 'solo');

    expect(ensureAccount).toHaveBeenCalledWith('ca@example.com');
    expect(createSubscription).toHaveBeenCalledWith(
      expect.objectContaining({
        notes: { billingEntityId: 'ca@example.com', planTier: 'solo', email: 'ca@example.com' },
      }),
    );
  });

  it('rejects a blank email', async () => {
    const { service } = buildService();
    await expect(service.createGuestSubscriptionSession('   ', 'solo')).rejects.toThrow(BadRequestException);
  });
});

describe('RazorpayCheckoutService.createTopupSession', () => {
  it('rejects an unknown pack id', async () => {
    const { service } = buildService();
    await expect(
      service.createTopupSession('user-1', 'ca@example.com', 'huge' as never),
    ).rejects.toThrow(BadRequestException);
  });

  it('ensures the account exists before creating the payment link', async () => {
    const { service, ensureAccount } = buildService();
    await service.createTopupSession('user-1', 'ca@example.com', 'medium');
    expect(ensureAccount).toHaveBeenCalledWith('user-1');
  });

  it('creates a one-time Payment Link priced and described from TOPUP_PACKS, with billingEntityId/packId/credits in notes', async () => {
    const { service, createPaymentLink } = buildService();
    await service.createTopupSession('user-1', 'ca@example.com', 'medium');

    expect(createPaymentLink).toHaveBeenCalledWith(
      expect.objectContaining({
        amount: 39900,
        currency: 'INR',
        description: '1000 Cellix credits',
        notes: { billingEntityId: 'user-1', packId: 'medium', credits: 1000 },
      }),
    );
  });

  it('returns the payment link short_url', async () => {
    const { service } = buildService();
    const result = await service.createTopupSession('user-1', 'ca@example.com', 'small');
    expect(result).toEqual({ url: 'https://rzp.io/i/link_x' });
  });

  it('throws when the payment link has no short_url', async () => {
    const { service } = buildService({ createPaymentLinkImpl: () => Promise.resolve({ short_url: undefined }) });
    await expect(service.createTopupSession('user-1', 'ca@example.com', 'small')).rejects.toThrow(
      ServiceUnavailableException,
    );
  });
});
