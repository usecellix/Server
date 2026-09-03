import { ModelRouter, COST_CAP_USD } from '../src/excel-ai/llm/model-router';
import { AppConfigService } from '../src/config/app-config.service';
import { ConversationTurn, MODEL_CONFIGS, WorkbookContext } from '../src/types/cellix.types';

/**
 * `ModelRouter.route()`'s HIGH -> MEDIUM downgrade-on-cost-cap had no dedicated
 * test anywhere in this repo — found while confirming spec 16 fix #4's
 * `estimateLlmCallCostUsd` refactor left `route()`'s behavior unchanged (it
 * does; `estimateCostUsd` now just delegates to it with identical inputs).
 * This file closes that pre-existing gap. Scope: only the downgrade behavior
 * and its recalculated cost — not a general ModelRouter/scoreTaskComplexity
 * test suite.
 */

function sheet(sheetName: string, rowCount: number, colCount: number): WorkbookContext['sheets'][number] {
  return {
    sheetName,
    usedRange: 'A1',
    rowCount,
    colCount,
    headers: [],
    sampleData: [],
  };
}

/** Scores 'high' via scoreTaskComplexity (verified below) independent of any
 * pricing input — cells(30) + formulas(9: vlookup/xlookup/sumif) +
 * sheets(20: 5 sheets) + cross-sheet(15: mentions "Sheet2") +
 * aggregation(10: "consolidate") = 84, which is > 65. */
function highComplexityContext(): WorkbookContext {
  return {
    activeSheet: 'Sheet1',
    sheets: [
      sheet('Sheet1', 100, 10),
      sheet('Sheet2', 100, 10),
      sheet('Sheet3', 100, 10),
      sheet('Sheet4', 100, 10),
      sheet('Sheet5', 100, 10),
    ],
  };
}

const HIGH_COMPLEXITY_PROMPT =
  'vlookup xlookup sumif consolidate Sheet2 cross-sheet operation across the workbook';

const NO_HISTORY: ConversationTurn[] = [];

function buildRouter(): ModelRouter {
  const config = {
    openRouterModelLow: 'openai/gpt-5-nano',
    openRouterModelMedium: 'openai/gpt-5-mini',
    openRouterModelHigh: 'openai/gpt-5',
  } as unknown as AppConfigService;
  return new ModelRouter(config);
}

describe('ModelRouter.route() — HIGH cost-cap downgrade', () => {
  it('fixture sanity check: scores HIGH complexity independent of promptTokenEstimate', () => {
    const router = buildRouter();
    // A tiny promptTokenEstimate keeps HIGH pricing itself under the cap here,
    // isolating that the tier-84 score is real and not an artifact of the
    // downgrade path below.
    const routing = router.route(HIGH_COMPLEXITY_PROMPT, highComplexityContext(), NO_HISTORY, 100);
    expect(routing.complexityScore.total).toBe(84);
    expect(routing.complexityScore.tier).toBe('high');
    expect(routing.tier).toBe('high');
    expect(routing.fallbackUsed).toBe(false);
  });

  it('downgrades high -> medium when estimatedCostUsd would exceed COST_CAP_USD, and recalculates cost against MEDIUM rates rather than reporting the pre-downgrade HIGH number', () => {
    const router = buildRouter();

    // HIGH:   (60000/1000)*0.00125 + (8192/1000)*0.01  = 0.075   + 0.08192  = 0.15692  (> COST_CAP_USD 0.15)
    // MEDIUM: (60000/1000)*0.00025 + (4096/1000)*0.002 = 0.015   + 0.008192 = 0.023192 (well under)
    const promptTokenEstimate = 60000;
    const expectedHighCost =
      (promptTokenEstimate / 1000) * MODEL_CONFIGS.high.costPer1kPrompt +
      (MODEL_CONFIGS.high.maxTokens / 1000) * MODEL_CONFIGS.high.costPer1kCompletion;
    const expectedMediumCost =
      (promptTokenEstimate / 1000) * MODEL_CONFIGS.medium.costPer1kPrompt +
      (MODEL_CONFIGS.medium.maxTokens / 1000) * MODEL_CONFIGS.medium.costPer1kCompletion;

    // Sanity on the fixture's premise: HIGH really does cross the cap here,
    // and MEDIUM really does not — otherwise the downgrade assertions below
    // could pass for the wrong reason.
    expect(expectedHighCost).toBeGreaterThan(COST_CAP_USD);
    expect(expectedMediumCost).toBeLessThan(COST_CAP_USD);

    const routing = router.route(HIGH_COMPLEXITY_PROMPT, highComplexityContext(), NO_HISTORY, promptTokenEstimate);

    // The complexity SCORE still says high — only the routing decision downgraded.
    expect(routing.complexityScore.tier).toBe('high');
    expect(routing.tier).toBe('medium');
    expect(routing.fallbackUsed).toBe(true);
    expect(routing.config.tier).toBe('medium');
    expect(routing.model).toBe('openai/gpt-5-mini');

    // The reported cost must be the RECALCULATED medium-tier figure ...
    expect(routing.estimatedCostUsd).toBeCloseTo(expectedMediumCost, 6);
    // ... not the pre-downgrade HIGH estimate that triggered the downgrade in
    // the first place (the exact bug this test guards against: reporting a
    // number priced at a model the request will not actually use).
    expect(routing.estimatedCostUsd).not.toBeCloseTo(expectedHighCost, 6);
    expect(routing.estimatedCostUsd).toBeLessThan(COST_CAP_USD);
  });

  it('does not downgrade when the HIGH-tier estimate stays under the cost cap', () => {
    const router = buildRouter();
    // Same high-complexity fixture, but a small promptTokenEstimate keeps HIGH
    // pricing comfortably under COST_CAP_USD — isolates that the downgrade is
    // conditioned on cost, not merely on tier === 'high'.
    const routing = router.route(HIGH_COMPLEXITY_PROMPT, highComplexityContext(), NO_HISTORY, 500);

    expect(routing.tier).toBe('high');
    expect(routing.fallbackUsed).toBe(false);
    expect(routing.model).toBe('openai/gpt-5');
    expect(routing.estimatedCostUsd).toBeLessThan(COST_CAP_USD);
  });
});
