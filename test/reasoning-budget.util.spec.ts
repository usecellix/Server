import {
  computeReasoningMaxTokens,
  ensureBudgetForReasoning,
  escalatedRetryBudget,
  isReasoningModel,
} from '../src/excel-ai/utils/reasoning-budget.util';

describe('isReasoningModel', () => {
  it.each(['openai/gpt-5', 'openai/gpt-5-mini', 'gpt-5', 'openai/o3', 'openai/o1'])(
    'detects %s as a reasoning model',
    (model) => {
      expect(isReasoningModel(model)).toBe(true);
    },
  );

  it.each(['openai/gpt-4o', 'openai/gpt-4o-mini', 'google/gemini-flash-1.5', undefined])(
    'treats %s as a non-reasoning model',
    (model) => {
      expect(isReasoningModel(model)).toBe(false);
    },
  );
});

describe('computeReasoningMaxTokens', () => {
  it('returns undefined for non-reasoning models so no cap is sent', () => {
    expect(computeReasoningMaxTokens('openai/gpt-4o', 4096)).toBeUndefined();
  });

  it('reserves room for content on the F5 incident budget', () => {
    // The real failure: gpt-5 at 4096 spent all 4096 reasoning, emitting nothing.
    const cap = computeReasoningMaxTokens('openai/gpt-5', 4096);
    expect(cap).toBeDefined();
    expect(cap!).toBeLessThan(4096);
    expect(4096 - cap!).toBeGreaterThanOrEqual(512);
  });

  it('never caps reasoning below the usable floor', () => {
    expect(computeReasoningMaxTokens('openai/gpt-5', 1024)).toBeGreaterThanOrEqual(512);
  });

  it('returns undefined when the budget is too small to split', () => {
    // Callers tuned for non-reasoning models pass 256/512; splitting is meaningless.
    expect(computeReasoningMaxTokens('openai/gpt-5', 256)).toBeUndefined();
  });

  it('guards against non-finite or non-positive budgets', () => {
    expect(computeReasoningMaxTokens('openai/gpt-5', 0)).toBeUndefined();
    expect(computeReasoningMaxTokens('openai/gpt-5', Number.NaN)).toBeUndefined();
  });
});

describe('ensureBudgetForReasoning', () => {
  it('leaves non-reasoning budgets untouched', () => {
    expect(ensureBudgetForReasoning('openai/gpt-4o', 256)).toBe(256);
  });

  it('raises starved budgets so both phases fit', () => {
    // llm-router passes 256, smart-data-query 512 — unusable for a reasoning model.
    expect(ensureBudgetForReasoning('openai/gpt-5', 256)).toBeGreaterThanOrEqual(1024);
    expect(ensureBudgetForReasoning('openai/gpt-5', 512)).toBeGreaterThanOrEqual(1024);
  });

  it('does not lower an already-generous budget', () => {
    expect(ensureBudgetForReasoning('openai/gpt-5', 8192)).toBe(8192);
  });
});

describe('escalatedRetryBudget', () => {
  it('always differs from the original budget', () => {
    // F5: the old retry recomputed the SAME 4096 and burned 40s for nothing.
    expect(escalatedRetryBudget(4096)).toBeGreaterThan(4096);
  });

  it('clamps to a sane ceiling', () => {
    expect(escalatedRetryBudget(100000)).toBeLessThanOrEqual(16384);
  });

  it('lifts tiny budgets to something workable', () => {
    expect(escalatedRetryBudget(256)).toBeGreaterThanOrEqual(2048);
  });
});
