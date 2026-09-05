import { CreditLedgerService } from '../src/credit/credit-ledger.service';
import { InsufficientCreditError } from '../src/credit/errors/insufficient-credit.error';

type Balances = { planCredits: number; purchasedCredits: number; oneTimeCredits: number };

/**
 * Mocks the atomic findOneAndUpdate against an in-memory balance snapshot,
 * applying the same fixed consumption order (planCredits, purchasedCredits,
 * oneTimeCredits) and floor check the real Mongo aggregation-pipeline
 * update performs, so these tests exercise the service's fallthrough logic
 * without a live database. CREDIT_SYSTEM.md CD-6 / CD-8.
 */
function buildService(initial: Balances | null) {
  let stored: Balances | null = initial ? { ...initial } : null;
  const insertedRows: Record<string, unknown>[] = [];
  const grants: { op: string; amount: number }[] = [];

  const findOne = jest.fn(() => ({
    lean: () => Promise.resolve(stored ? { ...stored } : null),
  }));

  const findOneAndUpdate = jest.fn((filter: { billingEntityId: string }, _pipeline: unknown) => {
    if (!stored) return Promise.resolve(null);
    const total = stored.planCredits + stored.purchasedCredits + stored.oneTimeCredits;
    const cost = lastCost;
    if (total < cost) return Promise.resolve(null);

    const planConsumed = Math.min(cost, stored.planCredits);
    const afterPlan = cost - planConsumed;
    const purchasedConsumed = Math.min(afterPlan, stored.purchasedCredits);
    const afterPurchased = afterPlan - purchasedConsumed;
    const oneTimeConsumed = Math.min(afterPurchased, stored.oneTimeCredits);

    stored = {
      planCredits: stored.planCredits - planConsumed,
      purchasedCredits: stored.purchasedCredits - purchasedConsumed,
      oneTimeCredits: stored.oneTimeCredits - oneTimeConsumed,
    };
    return Promise.resolve({ ...stored });
  });

  const updateOne = jest.fn((_filter: unknown, update: { $inc: Record<string, number> }) => {
    const [bucket, amount] = Object.entries(update.$inc)[0];
    if (stored) {
      (stored as unknown as Record<string, number>)[bucket] += amount;
    }
    grants.push({ op: bucket, amount });
    return Promise.resolve();
  });

  const insertMany = jest.fn((rows: Record<string, unknown>[]) => {
    insertedRows.push(...rows);
    return Promise.resolve();
  });

  const create = jest.fn((doc: Record<string, unknown>) => {
    insertedRows.push(doc);
    return Promise.resolve(doc);
  });

  let lastCost = 0;
  const creditAccountModel = { findOne, findOneAndUpdate, updateOne };
  const creditLedgerModel = { insertMany, create };
  const service = new CreditLedgerService(creditAccountModel as never, creditLedgerModel as never);

  return {
    service,
    insertedRows,
    grants,
    findOneAndUpdate,
    setLastCost: (cost: number) => {
      lastCost = cost;
    },
    getStored: () => stored,
  };
}

describe('CreditLedgerService.debit', () => {
  it('debits planCredits first, then purchasedCredits, then oneTimeCredits (CD-8)', async () => {
    const { service, insertedRows, setLastCost } = buildService({
      planCredits: 3,
      purchasedCredits: 2,
      oneTimeCredits: 10,
    });
    setLastCost(8); // FORMULA_GENERATE_OR_FIX

    const result = await service.debit('user-1', 'FORMULA_GENERATE_OR_FIX', 1, {
      conversationId: 'conv-1',
    });

    expect(result.debited).toBe(true);
    expect(result.balances).toEqual({ planCredits: 0, purchasedCredits: 0, oneTimeCredits: 7 });

    // 3 from plan, 2 from purchased, 3 from oneTime = 8 total, three rows.
    expect(insertedRows).toEqual([
      expect.objectContaining({ bucket: 'planCredits', amount: -3, actionType: 'FORMULA_GENERATE_OR_FIX' }),
      expect.objectContaining({ bucket: 'purchasedCredits', amount: -2 }),
      expect.objectContaining({ bucket: 'oneTimeCredits', amount: -3 }),
    ]);
  });

  it('returns debited: false without throwing when balance is insufficient', async () => {
    const { service, insertedRows, setLastCost } = buildService({
      planCredits: 1,
      purchasedCredits: 0,
      oneTimeCredits: 0,
    });
    setLastCost(8);

    const result = await service.debit('user-1', 'FORMULA_GENERATE_OR_FIX');

    expect(result).toEqual({ debited: false });
    expect(insertedRows).toEqual([]);
  });

  it('throws InsufficientCreditError for an unwired action type rather than silently debiting', async () => {
    const { service } = buildService({ planCredits: 1000, purchasedCredits: 0, oneTimeCredits: 0 });
    await expect(service.debit('user-1', 'GST_RECONCILIATION')).rejects.toThrow(InsufficientCreditError);
  });

  it('scales the debit by quantity for per-unit categories (CD-2)', async () => {
    const { service, setLastCost } = buildService({
      planCredits: 20,
      purchasedCredits: 0,
      oneTimeCredits: 0,
    });
    setLastCost(12); // ceil(340/100) * 3

    const result = await service.debit('user-1', 'GSTIN_VALIDATION_BATCH', 340);

    expect(result.debited).toBe(true);
    expect(result.balances?.planCredits).toBe(8);
  });

  it('records seatUserId on debit rows for pooled org accounts (CD-9)', async () => {
    const { service, insertedRows, setLastCost } = buildService({
      planCredits: 100,
      purchasedCredits: 0,
      oneTimeCredits: 0,
    });
    setLastCost(8);

    await service.debit('org-1', 'FORMULA_GENERATE_OR_FIX', 1, { seatUserId: 'user-42' });

    expect(insertedRows[0]).toEqual(expect.objectContaining({ seatUserId: 'user-42' }));
  });
});

describe('CreditLedgerService grants', () => {
  it('grantPlanCredits increments planCredits and writes a grant ledger row', async () => {
    const { service, insertedRows, grants } = buildService({
      planCredits: 0,
      purchasedCredits: 0,
      oneTimeCredits: 0,
    });

    await service.grantPlanCredits('user-1', 500, 'evt_123');

    expect(grants).toEqual([{ op: 'planCredits', amount: 500 }]);
    expect(insertedRows).toEqual([
      expect.objectContaining({ entryType: 'grant', amount: 500, bucket: 'planCredits', stripeEventId: 'evt_123' }),
    ]);
  });

  it('addPurchasedCredits increments purchasedCredits and writes a purchase ledger row', async () => {
    const { service, insertedRows, grants } = buildService({
      planCredits: 0,
      purchasedCredits: 0,
      oneTimeCredits: 0,
    });

    await service.addPurchasedCredits('user-1', 100);

    expect(grants).toEqual([{ op: 'purchasedCredits', amount: 100 }]);
    expect(insertedRows).toEqual([
      expect.objectContaining({ entryType: 'purchase', amount: 100, bucket: 'purchasedCredits' }),
    ]);
  });
});
