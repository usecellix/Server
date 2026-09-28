import { UsageBillingService } from '../src/credit/usage-billing.service';
import { LlmUsageService } from '../src/llm-usage/llm-usage.service';
import { runWithLlmUsageContext, type LlmUsageContext } from '../src/llm-usage/llm-usage.context';

function createBilling(debitUsage: jest.Mock, creditsPerUsd = 600) {
  const config = { creditsPerUsd } as never;
  const gate = { hasAnyBalance: jest.fn().mockResolvedValue(true) } as never;
  const ledger = { debitUsage } as never;
  return new UsageBillingService(config, gate, ledger);
}

const fullDebit = jest.fn(async (_user: string, credits: number) => ({
  debited: credits,
  requested: credits,
  balances: { planCredits: 0, purchasedCredits: 0, oneTimeCredits: 100 - credits },
}));

describe('UsageBillingService.creditsForCost', () => {
  const billing = createBilling(jest.fn());

  it('charges nothing for a request that made no model call', () => {
    expect(billing.creditsForCost(0)).toBe(0);
  });

  it('rounds real cost up to whole credits at CREDITS_PER_USD', () => {
    expect(billing.creditsForCost(0.0028)).toBe(2); // 1.68 credits
    expect(billing.creditsForCost(0.1)).toBe(60);
  });

  it('does not let float noise tip an exact amount up by a credit', () => {
    expect(billing.creditsForCost(0.1 + 0.2)).toBe(180);
  });
});

describe('UsageBillingService.settle', () => {
  beforeEach(() => fullDebit.mockClear());

  it('debits the request cost converted to credits, tagged with the prompt', async () => {
    const billing = createBilling(fullDebit);
    const context: LlmUsageContext = { promptId: 'p1', userId: 'u1', conversationId: 'c1', costUsd: 0.05 };

    const result = await billing.settle(context);

    expect(fullDebit).toHaveBeenCalledWith('u1', 30, { conversationId: 'c1', promptId: 'p1', costUsd: 0.05 });
    expect(result).toMatchObject({ debited: 30, requested: 30, costUsd: 0.05 });
  });

  it('bills only cost accrued since the previous settle', async () => {
    const billing = createBilling(fullDebit);
    const context: LlmUsageContext = { promptId: 'p1', userId: 'u1', costUsd: 0.05 };

    await billing.settle(context);
    expect(await billing.settle(context)).toBeNull();
    context.costUsd = 0.08;
    await billing.settle(context);

    expect(fullDebit).toHaveBeenCalledTimes(2);
    expect(fullDebit.mock.calls[1][1]).toBe(18); // 0.03 × 600
  });

  it('never double-charges when the stream-end and request-end settles overlap', async () => {
    const billing = createBilling(fullDebit);
    const context: LlmUsageContext = { promptId: 'p1', userId: 'u1', costUsd: 0.05 };

    await Promise.all([billing.settle(context), billing.settle(context)]);

    expect(fullDebit).toHaveBeenCalledTimes(1);
  });

  it('reports a short debit when the balance runs out, instead of refusing it', async () => {
    const partial = jest.fn().mockResolvedValue({
      debited: 20,
      requested: 60,
      balances: { planCredits: 0, purchasedCredits: 0, oneTimeCredits: 0 },
    });
    const billing = createBilling(partial);

    const result = await billing.settle({ promptId: 'p1', userId: 'u1', costUsd: 0.1 });

    expect(result).toMatchObject({ debited: 20, requested: 60 });
    expect(result?.balances?.oneTimeCredits).toBe(0);
  });

  it('skips anonymous requests and zero-cost requests', async () => {
    const billing = createBilling(fullDebit);

    expect(await billing.settle({ promptId: 'p1', costUsd: 0.1 })).toBeNull();
    expect(await billing.settle({ promptId: 'p1', userId: 'u1' })).toBeNull();
    expect(fullDebit).not.toHaveBeenCalled();
  });

  it('returns null rather than throwing when the ledger write fails', async () => {
    const billing = createBilling(jest.fn().mockRejectedValue(new Error('mongo down')));

    await expect(billing.settle({ promptId: 'p1', userId: 'u1', costUsd: 0.1 })).resolves.toBeNull();
  });
});

describe('LlmUsageService.recordCall feeds the per-request cost', () => {
  it('adds each call cost to the active request context, retries included', () => {
    const model = {
      create: jest.fn().mockResolvedValue(undefined),
      updateOne: jest.fn().mockReturnValue({ exec: jest.fn().mockResolvedValue(undefined) }),
    };
    const usage = new LlmUsageService(model as never, model as never);
    const context: LlmUsageContext = { promptId: 'p1', userId: 'u1' };
    const call = { model: 'm', caller: 'planner', streaming: false, costEstimated: false, latencyMs: 1, success: true };

    runWithLlmUsageContext(context, () => {
      usage.recordCall({ ...call, attempt: 1, costUsd: 0.01 });
      usage.recordCall({ ...call, attempt: 2, costUsd: 0.02 });
      usage.recordCall({ ...call, attempt: 1, costUsd: 0 });
    });

    expect(context.costUsd).toBeCloseTo(0.03, 10);
  });
});
