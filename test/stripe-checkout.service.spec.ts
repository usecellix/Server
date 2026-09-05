import { BadRequestException, ServiceUnavailableException } from '@nestjs/common';
import { StripeCheckoutService } from '../src/credit/stripe-checkout.service';

function buildService(options: {
  stripeSecretKey?: string;
  stripePriceSoloMonthly?: string;
  stripePriceFirmMonthly?: string;
  createSessionImpl?: (...args: unknown[]) => unknown;
} = {}) {
  const config = {
    stripeSecretKey: options.stripeSecretKey ?? 'sk_test_x',
    stripePriceSoloMonthly: options.stripePriceSoloMonthly ?? 'price_solo',
    stripePriceFirmMonthly: options.stripePriceFirmMonthly ?? 'price_firm',
    checkoutSuccessUrl: 'http://localhost:5173/checkout/success',
    checkoutCancelUrl: 'http://localhost:5173/checkout',
  };
  const ensureAccount = jest.fn().mockResolvedValue({});
  const creditGate = { ensureAccount };

  const service = new StripeCheckoutService(config as never, creditGate as never);

  const createSessions = jest.fn(
    options.createSessionImpl ?? (() => Promise.resolve({ url: 'https://checkout.stripe.com/session_x' })),
  );
  (service as unknown as { stripeClient?: { checkout: { sessions: { create: unknown } } } }).stripeClient = {
    checkout: { sessions: { create: createSessions } },
  } as never;

  return { service, ensureAccount, createSessions };
}

describe('StripeCheckoutService.createSubscriptionSession', () => {
  it('rejects a plan tier outside solo/firm', async () => {
    const { service } = buildService();
    await expect(
      service.createSubscriptionSession('user-1', 'ca@example.com', 'enterprise' as never),
    ).rejects.toThrow(BadRequestException);
  });

  it('ensures the account exists before creating the session', async () => {
    const { service, ensureAccount } = buildService();
    await service.createSubscriptionSession('user-1', 'ca@example.com', 'solo');
    expect(ensureAccount).toHaveBeenCalledWith('user-1');
  });

  it('creates a subscription-mode session with billingEntityId in client_reference_id and metadata', async () => {
    const { service, createSessions } = buildService();
    await service.createSubscriptionSession('user-1', 'ca@example.com', 'solo');

    expect(createSessions).toHaveBeenCalledWith(
      expect.objectContaining({
        mode: 'subscription',
        line_items: [{ price: 'price_solo', quantity: 1 }],
        client_reference_id: 'user-1',
        customer_email: 'ca@example.com',
        metadata: { billingEntityId: 'user-1', planTier: 'solo' },
      }),
    );
  });

  it('uses the firm price id for the firm tier', async () => {
    const { service, createSessions } = buildService();
    await service.createSubscriptionSession('org-1', undefined, 'firm');
    expect(createSessions).toHaveBeenCalledWith(
      expect.objectContaining({ line_items: [{ price: 'price_firm', quantity: 1 }] }),
    );
  });

  it('returns the session url', async () => {
    const { service } = buildService();
    const result = await service.createSubscriptionSession('user-1', 'ca@example.com', 'solo');
    expect(result).toEqual({ url: 'https://checkout.stripe.com/session_x' });
  });

  it('throws when the session has no url', async () => {
    const { service } = buildService({ createSessionImpl: () => Promise.resolve({ url: null }) });
    await expect(service.createSubscriptionSession('user-1', 'ca@example.com', 'solo')).rejects.toThrow(
      ServiceUnavailableException,
    );
  });

  it('throws when the price id for the plan tier is not configured', async () => {
    const { service } = buildService({ stripePriceSoloMonthly: '' });
    await expect(service.createSubscriptionSession('user-1', 'ca@example.com', 'solo')).rejects.toThrow(
      ServiceUnavailableException,
    );
  });
});

describe('StripeCheckoutService.createGuestSubscriptionSession', () => {
  it('normalizes the email (trim + lowercase) into billingEntityId', async () => {
    const { service, createSessions, ensureAccount } = buildService();
    await service.createGuestSubscriptionSession('  CA@Example.com  ', 'solo');

    expect(ensureAccount).toHaveBeenCalledWith('ca@example.com');
    expect(createSessions).toHaveBeenCalledWith(
      expect.objectContaining({
        client_reference_id: 'ca@example.com',
        customer_email: 'ca@example.com',
      }),
    );
  });

  it('rejects a blank email', async () => {
    const { service } = buildService();
    await expect(service.createGuestSubscriptionSession('   ', 'solo')).rejects.toThrow(BadRequestException);
  });
});
