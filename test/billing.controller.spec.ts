import { BadRequestException } from '@nestjs/common';
import { BillingController, PublicBillingController, RazorpayWebhookController } from '../src/credit/billing.controller';
import { AuthUserSession } from '../src/auth/auth.guard';

function session(userId: string, email = 'user@example.com'): AuthUserSession {
  return { user: { id: userId, email } } as unknown as AuthUserSession;
}

describe('BillingController', () => {
  it('GET /billing/account provisions the account (if needed) then returns the resolved summary', async () => {
    // Provision-on-read (Sept 9, 2026): a brand-new user's balance chip must
    // render on first look, not 404 until their first spend — see
    // getAccount's own docblock. ensureAccount is called unconditionally,
    // before the read; its own upsert-with-setOnInsert is what makes this
    // safe to call every time, not just on a genuinely-missing account.
    const getAccountSummary = jest.fn().mockResolvedValue({ availableBalance: 42 });
    const ensureAccount = jest.fn().mockResolvedValue(undefined);
    const controller = new BillingController(
      { getAccountSummary } as never,
      { ensureAccount } as never,
      {} as never,
    );

    const result = await controller.getAccount(session('user-1'));

    expect(ensureAccount).toHaveBeenCalledWith('user-1');
    expect(getAccountSummary).toHaveBeenCalledWith('user-1');
    expect(result).toEqual({ availableBalance: 42 });
  });

  it('GET /billing/ledger resolves billingEntityId from the session, never a request param', async () => {
    const getLedgerPage = jest.fn().mockResolvedValue({ entries: [], nextCursor: null });
    const controller = new BillingController({ getLedgerPage } as never, {} as never, {} as never);

    await controller.getLedger(session('user-1'), { limit: 10 });

    expect(getLedgerPage).toHaveBeenCalledWith('user-1', { limit: 10 });
  });

  it('POST /billing/checkout/subscribe resolves billingEntityId + email from the session, never the request body', async () => {
    const createSubscriptionSession = jest.fn().mockResolvedValue({ url: 'https://rzp.io/i/x' });
    const controller = new BillingController({} as never, {} as never, { createSubscriptionSession } as never);

    const result = await controller.createSubscribeCheckout(session('user-1', 'ca@example.com'), {
      planTier: 'solo',
    });

    expect(createSubscriptionSession).toHaveBeenCalledWith('user-1', 'ca@example.com', 'solo');
    expect(result).toEqual({ url: 'https://rzp.io/i/x' });
  });

  it('POST /billing/checkout/topup resolves billingEntityId + email from the session, never the request body', async () => {
    const createTopupSession = jest.fn().mockResolvedValue({ url: 'https://rzp.io/i/topup' });
    const controller = new BillingController({} as never, {} as never, { createTopupSession } as never);

    const result = await controller.createTopupCheckout(session('user-1', 'ca@example.com'), {
      packId: 'medium',
    });

    expect(createTopupSession).toHaveBeenCalledWith('user-1', 'ca@example.com', 'medium');
    expect(result).toEqual({ url: 'https://rzp.io/i/topup' });
  });
});

describe('PublicBillingController', () => {
  it('POST /billing/public/checkout/subscribe passes the submitted email through, unauthenticated', async () => {
    const createGuestSubscriptionSession = jest.fn().mockResolvedValue({ url: 'https://rzp.io/i/guest' });
    const controller = new PublicBillingController({ createGuestSubscriptionSession } as never);

    const result = await controller.createGuestSubscribeCheckout({
      email: 'ca@example.com',
      planTier: 'firm',
    });

    expect(createGuestSubscriptionSession).toHaveBeenCalledWith('ca@example.com', 'firm');
    expect(result).toEqual({ url: 'https://rzp.io/i/guest' });
  });
});

describe('RazorpayWebhookController', () => {
  function requestWithRawBody(rawBody: Buffer | undefined) {
    return { rawBody } as never;
  }

  it('verifies the raw body against the signature header and hands the parsed payload to handleVerifiedEvent', async () => {
    const fakePayload = { event: 'subscription.activated', payload: {} };
    const verifyAndParseEvent = jest.fn().mockReturnValue(fakePayload);
    const handleVerifiedEvent = jest.fn().mockResolvedValue({ alreadyProcessed: false });
    const controller = new RazorpayWebhookController({ verifyAndParseEvent, handleVerifiedEvent } as never);

    const rawBody = Buffer.from('{"event":"subscription.activated"}');
    const result = await controller.handleRazorpayWebhook(requestWithRawBody(rawBody), 'sig_abc');

    expect(verifyAndParseEvent).toHaveBeenCalledWith(rawBody, 'sig_abc');
    expect(handleVerifiedEvent).toHaveBeenCalledWith(fakePayload);
    expect(result).toEqual({ alreadyProcessed: false });
  });

  it('rejects when the raw body was not captured (content-type-parser hook missing)', async () => {
    const verifyAndParseEvent = jest.fn();
    const controller = new RazorpayWebhookController({ verifyAndParseEvent } as never);

    await expect(
      controller.handleRazorpayWebhook(requestWithRawBody(undefined), 'sig_abc'),
    ).rejects.toThrow(BadRequestException);
    expect(verifyAndParseEvent).not.toHaveBeenCalled();
  });
});
