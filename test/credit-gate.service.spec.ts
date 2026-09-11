import { CreditGateService, FREE_TIER_ONE_TIME_CREDITS } from '../src/credit/credit-gate.service';

/**
 * Mocks the upsert `ensureAccount` performs — findOneAndUpdate with
 * $setOnInsert — against an in-memory account, matching real Mongo's
 * behavior: an existing document is returned untouched ($setOnInsert only
 * applies on insert); a missing one is "inserted" as a Free-tier default.
 */
function buildService(account: { planCredits: number; purchasedCredits: number; oneTimeCredits: number } | null) {
  let stored = account;
  const findOneAndUpdate = jest.fn(() => ({
    lean: () => {
      if (!stored) {
        stored = { planCredits: 0, purchasedCredits: 0, oneTimeCredits: FREE_TIER_ONE_TIME_CREDITS };
      }
      return Promise.resolve(stored);
    },
  }));
  const creditAccountModel = { findOneAndUpdate };
  const service = new CreditGateService(creditAccountModel as never);
  return { service, findOneAndUpdate, getStored: () => stored };
}

/**
 * CREDIT_SYSTEM.md CD-4 — the gate check runs before dispatch, on the
 * *next* request; it never interrupts a run already in flight (that's
 * enforced elsewhere, by simply never being called mid-run).
 */
describe('CreditGateService.checkBalance', () => {
  it('allows a request when balance covers the flat catalog price', async () => {
    const { service } = buildService({ planCredits: 10, purchasedCredits: 0, oneTimeCredits: 0 });
    const result = await service.checkBalance('user-1', 'FORMULA_GENERATE_OR_FIX');
    expect(result).toEqual({ allowed: true, availableBalance: 10, requiredCredits: 8 });
  });

  it('blocks a request when balance is below the flat catalog price', async () => {
    const { service } = buildService({ planCredits: 3, purchasedCredits: 0, oneTimeCredits: 0 });
    const result = await service.checkBalance('user-1', 'FORMULA_GENERATE_OR_FIX');
    expect(result).toEqual({
      allowed: false,
      reason: 'insufficient_balance',
      availableBalance: 3,
      requiredCredits: 8,
    });
  });

  it('sums all three buckets toward the available balance', async () => {
    const { service } = buildService({ planCredits: 2, purchasedCredits: 3, oneTimeCredits: 3 });
    const result = await service.checkBalance('user-1', 'FORMULA_GENERATE_OR_FIX');
    expect(result).toEqual({ allowed: true, availableBalance: 8, requiredCredits: 8 });
  });

  it('lazily provisions a Free-tier account (30 oneTimeCredits) on first sight rather than treating it as zero balance', async () => {
    const { service, findOneAndUpdate } = buildService(null);
    const result = await service.checkBalance('user-missing', 'FORMULA_QA_SIMPLE');
    expect(result).toEqual({
      allowed: true,
      availableBalance: FREE_TIER_ONE_TIME_CREDITS,
      requiredCredits: 2,
    });
    expect(findOneAndUpdate).toHaveBeenCalledWith(
      { billingEntityId: 'user-missing' },
      expect.objectContaining({
        $setOnInsert: expect.objectContaining({
          billingEntityType: 'user',
          planTier: 'free',
          oneTimeCredits: FREE_TIER_ONE_TIME_CREDITS,
        }),
      }),
      expect.objectContaining({ upsert: true }),
    );
  });

  it('blocks once the lazily-provisioned Free-tier grant is exhausted', async () => {
    const { service } = buildService({ planCredits: 0, purchasedCredits: 0, oneTimeCredits: 1 });
    const result = await service.checkBalance('user-1', 'FORMULA_QA_COMPLEX');
    expect(result).toEqual({
      allowed: false,
      reason: 'insufficient_balance',
      availableBalance: 1,
      requiredCredits: 5,
    });
  });

  it('resolves per-unit cost with the requested quantity before checking (CD-2)', async () => {
    const { service } = buildService({ planCredits: 10, purchasedCredits: 0, oneTimeCredits: 0 });
    const result = await service.checkBalance('user-1', 'GSTIN_VALIDATION_BATCH', 340);
    expect(result).toEqual({
      allowed: false,
      reason: 'insufficient_balance',
      availableBalance: 10,
      requiredCredits: 12,
    });
  });

  it('blocks an unwired action type regardless of balance (CD-5)', async () => {
    const { service } = buildService({ planCredits: 1000, purchasedCredits: 0, oneTimeCredits: 0 });
    const result = await service.checkBalance('user-1', 'GST_RECONCILIATION');
    expect(result).toEqual({ allowed: false, reason: 'unwired_action' });
  });
});

describe('CreditGateService.hasAnyBalance', () => {
  it('is true when the account has a positive balance in any bucket', async () => {
    const { service } = buildService({ planCredits: 0, purchasedCredits: 0, oneTimeCredits: 1 });
    expect(await service.hasAnyBalance('user-1')).toBe(true);
  });

  it('is false when all three buckets are zero', async () => {
    const { service } = buildService({ planCredits: 0, purchasedCredits: 0, oneTimeCredits: 0 });
    expect(await service.hasAnyBalance('user-1')).toBe(false);
  });

  it('is true for a never-before-seen account, since provisioning grants 30 oneTimeCredits', async () => {
    const { service } = buildService(null);
    expect(await service.hasAnyBalance('user-missing')).toBe(true);
  });
});
