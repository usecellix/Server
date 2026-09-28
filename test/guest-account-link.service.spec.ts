import { GuestAccountLinkService } from '../src/credit/guest-account-link.service';

function buildService(options: {
  userByEmail?: Record<string, string>;
  guest?: { planCredits: number; purchasedCredits: number; planTier: string } | null;
} = {}) {
  const findUser = jest.fn((filter: { email: string }) => {
    const id = options.userByEmail?.[filter.email];
    return Promise.resolve(id ? { _id: id } : null);
  });
  const connection = { collection: jest.fn(() => ({ findOne: findUser })) };

  const accountUpdates: { filter: unknown; update: unknown }[] = [];
  const claimFilters: unknown[] = [];
  const creditAccountModel = {
    findOneAndUpdate: jest.fn((filter: unknown) => {
      claimFilters.push(filter);
      return { lean: () => Promise.resolve(options.guest ?? null) };
    }),
    updateOne: jest.fn((filter: unknown, update: unknown) => {
      accountUpdates.push({ filter, update });
      return Promise.resolve();
    }),
  };
  const ledgerRows: Record<string, unknown>[] = [];
  const creditLedgerModel = {
    insertMany: jest.fn((rows: Record<string, unknown>[]) => {
      ledgerRows.push(...rows);
      return Promise.resolve();
    }),
  };
  const subscriptionModel = { updateMany: jest.fn().mockResolvedValue(undefined) };

  const service = new GuestAccountLinkService(
    connection as never,
    creditAccountModel as never,
    creditLedgerModel as never,
    subscriptionModel as never,
  );
  return { service, accountUpdates, claimFilters, ledgerRows, subscriptionModel, findUser };
}

describe('GuestAccountLinkService.resolveBillingEntityId', () => {
  it('maps a guest email to the signed-in user with that email', async () => {
    const { service } = buildService({ userByEmail: { 'ca@example.com': 'user-1' } });
    await expect(service.resolveBillingEntityId('CA@Example.com')).resolves.toBe('user-1');
  });

  it('keeps the email when no user has signed up with it yet', async () => {
    const { service } = buildService();
    await expect(service.resolveBillingEntityId('new@example.com')).resolves.toBe('new@example.com');
  });

  it('passes a user id straight through without a lookup', async () => {
    const { service, findUser } = buildService();
    await expect(service.resolveBillingEntityId('6a5f4ed4821d0f40d58d8a1f')).resolves.toBe('6a5f4ed4821d0f40d58d8a1f');
    expect(findUser).not.toHaveBeenCalled();
  });
});

describe('GuestAccountLinkService.claimGuestAccount', () => {
  it('moves plan + purchased credits, the plan tier and subscriptions onto the user', async () => {
    const { service, accountUpdates, ledgerRows, subscriptionModel } = buildService({
      guest: { planCredits: 1000, purchasedCredits: 300, planTier: 'solo' },
    });

    await expect(service.claimGuestAccount('user-1', 'CA@example.com')).resolves.toBe(1300);

    expect(accountUpdates).toEqual([
      {
        filter: { billingEntityId: 'user-1' },
        update: { $inc: { planCredits: 1000, purchasedCredits: 300 }, $set: { planTier: 'solo' } },
      },
    ]);
    expect(subscriptionModel.updateMany).toHaveBeenCalledWith(
      { billingEntityId: 'ca@example.com' },
      { $set: { billingEntityId: 'user-1' } },
    );
    // Every credit moved shows up on both sides of the ledger.
    expect(ledgerRows.reduce((sum, r) => sum + (r.amount as number), 0)).toBe(0);
  });

  it('does nothing when there is no unclaimed guest account (atomic claim lost or never existed)', async () => {
    const { service, accountUpdates, ledgerRows } = buildService({ guest: null });

    await expect(service.claimGuestAccount('user-1', 'ca@example.com')).resolves.toBe(0);

    expect(accountUpdates).toEqual([]);
    expect(ledgerRows).toEqual([]);
  });

  it('only claims accounts not already linked', async () => {
    const { service, claimFilters } = buildService();
    await service.claimGuestAccount('user-1', 'ca@example.com');
    expect(claimFilters[0]).toEqual(expect.objectContaining({ billingEntityId: 'ca@example.com', linkedToUserId: { $exists: false } }));
  });

  it('skips a session with no email', async () => {
    const { service, claimFilters } = buildService();
    await expect(service.claimGuestAccount('user-1', undefined)).resolves.toBe(0);
    expect(claimFilters).toEqual([]);
  });
});
