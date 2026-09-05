import { CreditAccountQueryService } from '../src/credit/credit-account-query.service';

function buildService(account: Record<string, unknown> | null, ledgerDocs: Record<string, unknown>[]) {
  const accountFindOne = jest.fn(() => ({ lean: () => Promise.resolve(account) }));
  const creditAccountModel = { findOne: accountFindOne };

  const find = jest.fn(() => ({
    sort: () => ({
      limit: () => ({
        select: () => ({ lean: () => Promise.resolve(ledgerDocs) }),
      }),
    }),
  }));
  const creditLedgerModel = { find };

  const service = new CreditAccountQueryService(creditAccountModel as never, creditLedgerModel as never);
  return { service, accountFindOne, find };
}

describe('CreditAccountQueryService.getAccountSummary', () => {
  it('returns the summed balance and plan info for an existing account', async () => {
    const { service } = buildService(
      {
        billingEntityType: 'user',
        planTier: 'solo',
        planCredits: 400,
        purchasedCredits: 50,
        oneTimeCredits: 0,
        currentPeriodEnd: new Date('2026-10-01'),
      },
      [],
    );

    const summary = await service.getAccountSummary('user-1');

    expect(summary).toEqual({
      billingEntityType: 'user',
      planTier: 'solo',
      planCredits: 400,
      purchasedCredits: 50,
      oneTimeCredits: 0,
      availableBalance: 450,
      currentPeriodEnd: new Date('2026-10-01'),
    });
  });

  it('returns null (not a zero-balance object) when no account exists yet', async () => {
    const { service } = buildService(null, []);
    expect(await service.getAccountSummary('user-missing')).toBeNull();
  });

  it('does not provision an account as a side effect of reading it', async () => {
    // Unlike CreditGateService, this is a pure read path — no $setOnInsert
    // upsert should ever be issued here.
    const { service, accountFindOne } = buildService(null, []);
    await service.getAccountSummary('user-missing');
    expect(accountFindOne).toHaveBeenCalledWith({ billingEntityId: 'user-missing' });
  });
});

describe('CreditAccountQueryService.getLedgerPage', () => {
  it('returns entries and a null cursor when everything fits on one page', async () => {
    const docs = [
      { entryType: 'debit', amount: -8, bucket: 'planCredits', actionType: 'FORMULA_GENERATE_OR_FIX', createdAt: new Date('2026-09-01') },
    ];
    const { service } = buildService({}, docs);

    const page = await service.getLedgerPage('user-1', { limit: 25 });

    expect(page.entries).toHaveLength(1);
    expect(page.nextCursor).toBeNull();
  });

  it('sets nextCursor to the last row of the page when there is one more beyond it', async () => {
    // limit+1 fetch pattern: 3 docs returned for a limit of 2 means "more exist".
    const docs = [
      { entryType: 'debit', amount: -8, bucket: 'planCredits', createdAt: new Date('2026-09-03') },
      { entryType: 'debit', amount: -2, bucket: 'planCredits', createdAt: new Date('2026-09-02') },
      { entryType: 'debit', amount: -5, bucket: 'planCredits', createdAt: new Date('2026-09-01') },
    ];
    const { service } = buildService({}, docs);

    const page = await service.getLedgerPage('user-1', { limit: 2 });

    expect(page.entries).toHaveLength(2);
    expect(page.nextCursor).toBe(new Date('2026-09-02').toISOString());
  });

  it('rejects a malformed cursor', async () => {
    const { service } = buildService({}, []);
    await expect(service.getLedgerPage('user-1', { cursor: 'not-a-date' })).rejects.toThrow();
  });
});
