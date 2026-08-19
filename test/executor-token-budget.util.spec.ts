import {
  EXECUTOR_BASE_MAX_TOKENS,
  EXECUTOR_MAX_TOKENS_CAP,
  resolveExecutorMaxTokens,
} from '../src/agents/utils/executor-token-budget.util';

/**
 * Regression: the Executor's completion budget was a flat 2000 tokens
 * regardless of subtask size — a formula table spanning many sheets (e.g. one
 * row per month, several SUMIF formulas per row) could truncate mid-way, and
 * because every retry reused the same fixed budget, it failed the exact same
 * way twice ("failed after 2 attempts, formulas only through April").
 */
describe('resolveExecutorMaxTokens', () => {
  it('uses the base budget for a small (1-2 action) subtask', () => {
    expect(resolveExecutorMaxTokens(1)).toBe(EXECUTOR_BASE_MAX_TOKENS + 60);
    expect(resolveExecutorMaxTokens(2)).toBe(EXECUTOR_BASE_MAX_TOKENS + 120);
  });

  it('scales up for a subtask estimating many actions (e.g. a multi-sheet formula table)', () => {
    // ~18 formula cells, matching the Jan-Jun half of the chunked monthly totals table.
    const budget = resolveExecutorMaxTokens(18);
    expect(budget).toBeGreaterThan(EXECUTOR_BASE_MAX_TOKENS);
    expect(budget).toBe(EXECUTOR_BASE_MAX_TOKENS + 18 * 60);
  });

  it('defaults to the base budget when estimatedActions is missing', () => {
    expect(resolveExecutorMaxTokens(undefined)).toBe(EXECUTOR_BASE_MAX_TOKENS + 60);
  });

  it('never returns less than the base budget for a zero or negative estimate', () => {
    expect(resolveExecutorMaxTokens(0)).toBe(EXECUTOR_BASE_MAX_TOKENS + 60);
    expect(resolveExecutorMaxTokens(-5)).toBe(EXECUTOR_BASE_MAX_TOKENS + 60);
  });

  it('caps the budget so a runaway estimate cannot exceed the provider limit', () => {
    expect(resolveExecutorMaxTokens(10_000)).toBe(EXECUTOR_MAX_TOKENS_CAP);
  });

  it('floors a fractional estimate before scaling', () => {
    expect(resolveExecutorMaxTokens(3.9)).toBe(EXECUTOR_BASE_MAX_TOKENS + 3 * 60);
  });
});
