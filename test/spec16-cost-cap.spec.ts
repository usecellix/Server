import { PlannerAgent } from '../src/agents/planner.agent';
import {
  PLANNER_COST_CAP_USER_MESSAGE,
  PlannerCostCapExceededError,
  PlannerExhaustedError,
} from '../src/agents/errors';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { WorkbookContext } from '../src/agents/types/agent.types';
import {
  COST_CAP_USD,
  estimateLlmCallCostUsd,
  resolvePricingForModel,
} from '../src/excel-ai/llm/model-router';
import { MODEL_CONFIGS } from '../src/types/cellix.types';

const DASHBOARD_PROMPT =
  'In dashboard create a chart ,and analysis for purchase register a summary for purchase register';

function sheet(
  name: string,
  values: unknown[][],
  rowCount = values.length,
  columnCount = values[0]?.length ?? 0,
): WorkbookContext['sheets'][number] {
  return {
    name,
    usedRange: rowCount && columnCount ? `A1:${String.fromCharCode(64 + columnCount)}${rowCount}` : 'A1',
    rowCount,
    columnCount,
    values,
    formulas: values.map((row) => row.map(() => '')),
    numberFormats: values.map((row) => row.map(() => 'General')),
    structure: 'data_table',
    headerRowIndex: 0,
  };
}

function nonEmptyContext(): WorkbookContext {
  return {
    activeSheetName: 'Dashboard',
    sheets: [
      sheet('Dashboard', [
        ['Col A', 'Col B'],
        [1, 2],
      ]),
      sheet('Purchase Register', [
        ['Item', 'Amount'],
        ['Widget', 100],
      ]),
    ],
    namedRanges: [],
    tables: [],
  };
}

function buildAgent(completeImpl: jest.Mock, configOverrides: Partial<AppConfigService> = {}): PlannerAgent {
  const llm = { complete: completeImpl } as unknown as OpenRouterService;
  const config = {
    openRouterModelLow: 'openai/gpt-5-mini',
    openRouterModelMedium: 'openai/gpt-5-mini',
    openRouterModelHigh: 'openai/gpt-5',
    openRouterModelPlanner: 'openai/gpt-5',
    ...configOverrides,
  } as unknown as AppConfigService;
  return new PlannerAgent(llm, config);
}

const goodPlan = JSON.stringify({
  subtasks: [
    {
      id: 's1',
      description: 'Create chart on Dashboard',
      targetSheet: 'Dashboard',
      dependsOn: [],
      estimatedActions: 2,
    },
  ],
  clarificationsNeeded: [],
  confidence: 'high',
  reasoning: 'ok',
});

describe('Spec 16 fix #4 — Tier-3 Planner cost cap', () => {
  it('estimateLlmCallCostUsd / resolvePricingForModel: a real HIGH-tier request is comfortably under cap', () => {
    const { pricing, approximate } = resolvePricingForModel('openai/gpt-5', {
      openRouterModelLow: 'openai/gpt-5-mini',
      openRouterModelMedium: 'openai/gpt-5-mini',
      openRouterModelHigh: 'openai/gpt-5',
    });
    expect(approximate).toBe(false);
    // Realistic dashboard-sized prompt (a few thousand tokens) at the last-resort ceiling.
    const cost = estimateLlmCallCostUsd(pricing, 3000, 8192);
    expect(cost).toBeLessThan(COST_CAP_USD);
  });

  it('resolvePricingForModel falls back to HIGH pricing (not a silent under-price) and flags approximate for an unrecognized model', () => {
    const result = resolvePricingForModel('some-provider/custom-eval-model', {
      // LOW and MEDIUM deliberately distinct here so a wrong fallback to either
      // one would be caught by the exact-equality check below, not just masked
      // by them happening to share a rate with HIGH.
      openRouterModelLow: 'openai/gpt-5-nano',
      openRouterModelMedium: 'openai/gpt-5-mini',
      openRouterModelHigh: 'openai/gpt-5',
    });
    expect(result.approximate).toBe(true);
    expect(result.tier).toBe('high');
    // Exact equality, not a shape check — proves this is genuinely HIGH's rate
    // ($0.00125/$0.01 per 1k) and not LOW's ($0.00005/$0.0004) or MEDIUM's
    // ($0.00025/$0.002) silently passed through as the fallback.
    expect(result.pricing).toEqual(MODEL_CONFIGS.high);
    expect(result.pricing).not.toEqual(MODEL_CONFIGS.low);
    expect(result.pricing).not.toEqual(MODEL_CONFIGS.medium);
  });

  it('allows a normal-size compound request through unaffected — llm.complete is still called', async () => {
    const complete = jest.fn().mockImplementation((opts: { outcome?: { truncated?: boolean } }) => {
      if (opts.outcome) opts.outcome.truncated = false;
      return Promise.resolve(goodPlan);
    });
    const agent = buildAgent(complete);

    const plan = await agent.plan(
      DASHBOARD_PROMPT,
      nonEmptyContext(),
      [],
      undefined,
      'corr_cost_ok',
      undefined,
      3,
    );

    expect(complete).toHaveBeenCalledTimes(1);
    expect(plan.subtasks.length).toBeGreaterThan(0);
  });

  it('refuses with PlannerCostCapExceededError — not PlannerExhaustedError — before making any LLM call, for an oversized context', async () => {
    const complete = jest.fn().mockResolvedValue(goodPlan);
    const agent = buildAgent(complete);

    // A promptContext this large (~1MB) is embedded near-verbatim by
    // buildPlannerUserMessage, pushing the worst-case cost estimate
    // (promptTokens ~ length/4, priced at HIGH, plus the 8192-token
    // last-resort completion ceiling) well past COST_CAP_USD.
    const hugePromptContext = 'X'.repeat(1_000_000);

    let caught: unknown;
    try {
      await agent.plan(
        DASHBOARD_PROMPT,
        nonEmptyContext(),
        [],
        hugePromptContext,
        'corr_cost_exceeded',
        undefined,
        3,
      );
    } catch (error) {
      caught = error;
    }

    expect(caught).toBeInstanceOf(PlannerCostCapExceededError);
    expect(caught).not.toBeInstanceOf(PlannerExhaustedError);
    expect((caught as Error).message).toBe(PLANNER_COST_CAP_USER_MESSAGE);
    expect((caught as PlannerCostCapExceededError).estimatedCostUsd).toBeGreaterThan(COST_CAP_USD);
    expect((caught as PlannerCostCapExceededError).costCapUsd).toBe(COST_CAP_USD);

    // The whole point: refuse BEFORE spending anything on the call, not after.
    expect(complete).not.toHaveBeenCalled();
  });

  it('logs an APPROXIMATE-pricing warning when openRouterModelPlanner matches none of LOW/MEDIUM/HIGH', async () => {
    const complete = jest.fn().mockImplementation((opts: { outcome?: { truncated?: boolean } }) => {
      if (opts.outcome) opts.outcome.truncated = false;
      return Promise.resolve(goodPlan);
    });
    const agent = buildAgent(complete, {
      openRouterModelPlanner: 'some-provider/custom-eval-model',
    });
    const warnSpy = jest
      .spyOn((agent as unknown as { logger: { warn: (m: string) => void } }).logger, 'warn')
      .mockImplementation(() => undefined);

    await agent.plan(DASHBOARD_PROMPT, nonEmptyContext(), [], undefined, 'corr_approx', undefined, 3);

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining('APPROXIMATE'));
    warnSpy.mockRestore();
  });
});
