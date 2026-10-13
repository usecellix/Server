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

  /** Wraps a result so callers may either `await` it directly or chain `.lean()` first — matches real Mongoose. */
  function queryResult<T>(value: T) {
    return Object.assign(Promise.resolve(value), { lean: () => Promise.resolve(value) });
  }

  const findOneAndUpdate = jest.fn(
    (_filter: { billingEntityId: string }, update: unknown, options?: { new?: boolean; upsert?: boolean }) => {
      // grantPlanCredits: a plain `{ $set: { planCredits } }` reset, not the
      // debit/debitUsage aggregation pipeline (an array) below.
      if (!Array.isArray(update)) {
        const before = stored ? { ...stored } : { planCredits: 0, purchasedCredits: 0, oneTimeCredits: 0 };
        const set = (update as { $set?: Record<string, number> }).$set ?? {};
        stored = { ...before, ...set };
        return queryResult(options?.new === false ? before : { ...stored });
      }

      if (!stored) return queryResult(null);
      const total = stored.planCredits + stored.purchasedCredits + stored.oneTimeCredits;
      const cost = lastCost;
      if (total < cost) return queryResult(null);

      const planConsumed = Math.min(cost, stored.planCredits);
      const afterPlan = cost - planConsumed;
      const purchasedConsumed = Math.min(afterPlan, stored.purchasedCredits);
      const afterPurchased = afterPlan - purchasedConsumed;
      const oneTimeConsumed = Math.min(afterPurchased, stored.oneTimeCredits);

      const before = { ...stored };
      stored = {
        planCredits: stored.planCredits - planConsumed,
        purchasedCredits: stored.purchasedCredits - purchasedConsumed,
        oneTimeCredits: stored.oneTimeCredits - oneTimeConsumed,
      };
      return queryResult(options?.new === false ? before : { ...stored });
    },
  );

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

  // Enforces credit_ledger's unique paymentEventId index like Mongo would.
  const create = jest.fn((doc: Record<string, unknown>) => {
    if (doc.paymentEventId && insertedRows.some((row) => row.paymentEventId === doc.paymentEventId)) {
      return Promise.reject(Object.assign(new Error('E11000 duplicate key'), { code: 11000 }));
    }
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
  it('grantPlanCredits sets planCredits to the plan allotment from zero and writes a grant row, no expire row', async () => {
    const { service, insertedRows, getStored } = buildService({
      planCredits: 0,
      purchasedCredits: 0,
      oneTimeCredits: 0,
    });

    await service.grantPlanCredits('user-1', 500, 'evt_123');

    expect(getStored()).toEqual({ planCredits: 500, purchasedCredits: 0, oneTimeCredits: 0 });
    expect(insertedRows).toEqual([
      expect.objectContaining({ entryType: 'grant', amount: 500, bucket: 'planCredits', paymentEventId: 'evt_123' }),
    ]);
  });

  it('RESETS planCredits on renewal instead of stacking (TASKS.md #342) — leftover is logged as expired, not compounded', async () => {
    const { service, insertedRows, getStored } = buildService({
      planCredits: 1800, // unused leftover from a prior cycle
      purchasedCredits: 0,
      oneTimeCredits: 0,
    });

    await service.grantPlanCredits('user-1', 3000, 'evt_renewal_1');

    // Not 4800 — a renewal must not let unused credits compound indefinitely.
    expect(getStored()).toEqual({ planCredits: 3000, purchasedCredits: 0, oneTimeCredits: 0 });
    expect(insertedRows).toEqual([
      expect.objectContaining({ entryType: 'grant', amount: 3000, bucket: 'planCredits', paymentEventId: 'evt_renewal_1' }),
      expect.objectContaining({ entryType: 'expire', amount: -1800, bucket: 'planCredits' }),
    ]);
    // paymentEventId is a unique index — the real event id belongs on the
    // grant row only, never duplicated onto the paired expire row.
    expect(insertedRows[1].paymentEventId).toBeUndefined();
  });

  it('grantPlanCredits provisions the account when it does not exist yet (upsert)', async () => {
    const { service, getStored } = buildService(null);

    await service.grantPlanCredits('user-new', 3000, 'evt_1');

    expect(getStored()).toEqual({ planCredits: 3000, purchasedCredits: 0, oneTimeCredits: 0 });
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

  it('addPurchasedCredits grants once for a repeated paymentEventId (webhook + reconcile, or a redelivery)', async () => {
    const { service, grants, getStored } = buildService({ planCredits: 0, purchasedCredits: 0, oneTimeCredits: 0 });

    await expect(service.addPurchasedCredits('user-1', 300, 'topup:plink_1')).resolves.toBe(true);
    await expect(service.addPurchasedCredits('user-1', 300, 'topup:plink_1')).resolves.toBe(false);

    expect(grants).toHaveLength(1);
    expect(getStored()?.purchasedCredits).toBe(300);
  });

  it('grantPlanCredits does not reset a partly-spent balance when the same cycle is granted again', async () => {
    const { service, getStored, setLastCost } = buildService({ planCredits: 0, purchasedCredits: 0, oneTimeCredits: 0 });

    await service.grantPlanCredits('user-1', 3000, 'plan:sub_1:1700000000');
    setLastCost(8);
    await service.debit('user-1', 'FORMULA_GENERATE_OR_FIX');
    await expect(service.grantPlanCredits('user-1', 3000, 'plan:sub_1:1700000000')).resolves.toBe(false);

    expect(getStored()?.planCredits).toBe(2992);
  });
});
