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
 *
 * 2026-09-05: LOW/MEDIUM/HIGH were collapsed to a single model (openai/gpt-5)
 * across the board (TASKS.md #183). 2026-09-10: re-priced again for the GLM
 * model swap — MEDIUM and HIGH both now genuinely resolve to `z-ai/glm-5.3`
 * (MEDIUM directly, HIGH via the `glm-latest` alias), so they are still
 * byte-identical in `MODEL_CONFIGS`, just for an honest reason this time (the
 * right model for both tiers) rather than because pricing was never updated
 * after a swap. The downgrade-is-a-no-op assertions below still hold; only
 * the model name and the prompt-size needed to cross COST_CAP_USD changed,
 * since GLM pricing is ~18.7x cheaper per token than the gpt-5 figures this
 * test was originally tuned against.
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
    openRouterModelLow: 'z-ai/glm-5.3-flash',
    openRouterModelMedium: 'z-ai/glm-5.3',
    openRouterModelHigh: 'z-ai/glm-latest',
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

  it('still flips tier/fallbackUsed on cost-cap overage, but MEDIUM no longer resolves to a cheaper model or a lower dollar figure', () => {
    const router = buildRouter();

    // HIGH and MEDIUM are now byte-identical configs (both z-ai/glm-5.3) —
    // MEDIUM resolves to it directly, HIGH via the glm-latest alias — so
    // there is no cheaper fallback the downgrade can land on. GLM pricing is
    // ~18.7x cheaper per token than the gpt-5 pricing this test was
    // originally tuned against, so a much larger prompt is needed to cross
    // COST_CAP_USD under the new pricing.
    const promptTokenEstimate = 800000;
    const expectedCost =
      (promptTokenEstimate / 1000) * MODEL_CONFIGS.high.costPer1kPrompt +
      (MODEL_CONFIGS.high.maxTokens / 1000) * MODEL_CONFIGS.high.costPer1kCompletion;

    expect(MODEL_CONFIGS.medium.model).toBe(MODEL_CONFIGS.high.model);
    expect(MODEL_CONFIGS.medium.maxTokens).toBe(MODEL_CONFIGS.high.maxTokens);
    expect(MODEL_CONFIGS.medium.costPer1kPrompt).toBe(MODEL_CONFIGS.high.costPer1kPrompt);
    expect(MODEL_CONFIGS.medium.costPer1kCompletion).toBe(MODEL_CONFIGS.high.costPer1kCompletion);
    expect(expectedCost).toBeGreaterThan(COST_CAP_USD);

    const routing = router.route(HIGH_COMPLEXITY_PROMPT, highComplexityContext(), NO_HISTORY, promptTokenEstimate);

    // The complexity SCORE still says high, and the downgrade still fires —
    // model-router.ts's cap-check logic is untouched by the model repricing.
    expect(routing.complexityScore.tier).toBe('high');
    expect(routing.tier).toBe('medium');
    expect(routing.fallbackUsed).toBe(true);
    expect(routing.config.tier).toBe('medium');

    // ...but the "downgrade" no longer buys anything: same model, same price,
    // and the recalculated estimate still exceeds the cap it was meant to
    // enforce. This is the honest state of MEDIUM/HIGH sharing a model, not a
    // bug this test should paper over.
    expect(routing.model).toBe('z-ai/glm-5.3');
    expect(routing.estimatedCostUsd).toBeCloseTo(expectedCost, 6);
    expect(routing.estimatedCostUsd).toBeGreaterThan(COST_CAP_USD);
  });

  it('does not downgrade when the HIGH-tier estimate stays under the cost cap', () => {
    const router = buildRouter();
    // Same high-complexity fixture, but a small promptTokenEstimate keeps HIGH
    // pricing comfortably under COST_CAP_USD — isolates that the downgrade is
    // conditioned on cost, not merely on tier === 'high'.
    const routing = router.route(HIGH_COMPLEXITY_PROMPT, highComplexityContext(), NO_HISTORY, 500);

    expect(routing.tier).toBe('high');
    expect(routing.fallbackUsed).toBe(false);
    expect(routing.model).toBe('z-ai/glm-latest');
    expect(routing.estimatedCostUsd).toBeLessThan(COST_CAP_USD);
  });
});
