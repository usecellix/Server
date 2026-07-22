import { PlannerAgent } from '../src/agents/planner.agent';
import {
  PLANNER_EXHAUSTED_USER_MESSAGE,
  PlannerExhaustedError,
} from '../src/agents/errors';
import {
  PLANNER_LAST_RESORT_MAX_TOKENS,
  resolvePlannerMaxTokens,
} from '../src/agents/utils/planner-token-budget.util';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { WorkbookContext } from '../src/agents/types/agent.types';
import { ModelRouter } from '../src/excel-ai/llm/model-router';

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

describe('resolvePlannerMaxTokens', () => {
  it('gives Tier 3 / default a high ceiling for compound/dashboard work', () => {
    expect(resolvePlannerMaxTokens(3)).toBeGreaterThanOrEqual(3000);
    expect(resolvePlannerMaxTokens(undefined)).toBeGreaterThanOrEqual(3000);
  });

  it('uses smaller budgets for Tier 0–1', () => {
    expect(resolvePlannerMaxTokens(0)).toBeLessThan(resolvePlannerMaxTokens(3));
    expect(resolvePlannerMaxTokens(1)).toBeLessThan(resolvePlannerMaxTokens(3));
  });
});

describe('PlannerAgent token exhaustion (Spec 16)', () => {
  function buildAgent(completeImpl: jest.Mock): PlannerAgent {
    const llm = { complete: completeImpl } as unknown as OpenRouterService;
    const config = {
      openRouterModelHigh: 'openai/gpt-5',
    } as unknown as AppConfigService;
    return new PlannerAgent(llm, config);
  }

  it('passes tier-3 maxTokens >= 3000 on planner calls', async () => {
    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({
        subtasks: [
          {
            id: 's1',
            description: 'Create chart on Dashboard from Purchase Register',
            targetSheet: 'Dashboard',
            dependsOn: [],
            estimatedActions: 2,
          },
        ],
        clarificationsNeeded: [],
        confidence: 'high',
        reasoning: 'ok',
      }),
    );
    const agent = buildAgent(complete);

    await agent.plan(DASHBOARD_PROMPT, nonEmptyContext(), [], undefined, 'corr_1', undefined, 3);

    expect(complete).toHaveBeenCalled();
    const firstCall = complete.mock.calls[0][0] as { maxTokens: number; reasoningMaxTokens?: number };
    expect(firstCall.maxTokens).toBeGreaterThanOrEqual(3000);
    expect(firstCall.reasoningMaxTokens).toBeDefined();
  });

  it('throws PlannerExhaustedError instead of a stub single-step plan after failures', async () => {
    const complete = jest.fn().mockResolvedValue('');
    const agent = buildAgent(complete);

    await expect(
      agent.plan(DASHBOARD_PROMPT, nonEmptyContext(), [], undefined, 'corr_2', undefined, 3),
    ).rejects.toBeInstanceOf(PlannerExhaustedError);

    await expect(
      agent.plan(DASHBOARD_PROMPT, nonEmptyContext(), [], undefined, 'corr_2b', undefined, 3),
    ).rejects.toMatchObject({
      message: PLANNER_EXHAUSTED_USER_MESSAGE,
      originalMessage: DASHBOARD_PROMPT,
    });

    // First attempt + parse retry + last-resort
    expect(complete.mock.calls.length).toBeGreaterThanOrEqual(3);
    const lastCall = complete.mock.calls[complete.mock.calls.length - 1][0] as {
      maxTokens: number;
    };
    expect(lastCall.maxTokens).toBe(PLANNER_LAST_RESORT_MAX_TOKENS);
  });

  it('never returns reasoning "Fallback single-step plan — planner JSON was not parseable"', async () => {
    const complete = jest.fn().mockResolvedValue('not json at all {{{');
    const agent = buildAgent(complete);

    try {
      await agent.plan(
        'make a fancy dashboard with charts and summaries',
        nonEmptyContext(),
        [],
        undefined,
        'corr_3',
        undefined,
        3,
      );
      fail('expected PlannerExhaustedError');
    } catch (error) {
      expect(error).toBeInstanceOf(PlannerExhaustedError);
      expect(String((error as Error).message)).not.toContain(
        'Fallback single-step plan',
      );
    }
  });

  it('still returns empty-workbook clarification fallback when useful', async () => {
    const complete = jest.fn().mockResolvedValue('');
    const agent = buildAgent(complete);
    const emptyCtx: WorkbookContext = {
      activeSheetName: 'Sheet1',
      sheets: [sheet('Sheet1', [], 0, 0)],
      namedRanges: [],
      tables: [],
    };

    const plan = await agent.plan('sort by amount', emptyCtx, [], undefined, 'corr_4', undefined, 1);
    expect(plan.clarificationsNeeded.length).toBeGreaterThan(0);
    expect(plan.subtasks).toHaveLength(0);
  });
});

describe('OpenRouterService reasoning_token_exhaustion alert', () => {
  function buildService(): OpenRouterService {
    const config = {
      openRouterApiKey: 'test-key',
      openRouterHttpReferer: 'http://localhost',
      openRouterModelLow: 'openai/gpt-5-mini',
      openRouterModelMedium: 'openai/gpt-5-mini',
      openRouterModelHigh: 'openai/gpt-5',
    } as unknown as AppConfigService;
    const modelRouter = { markRateLimited: jest.fn() } as unknown as ModelRouter;
    return new OpenRouterService(config, modelRouter);
  }

  it('logs ALERT reasoning_token_exhaustion when completionTokens === reasoningTokens', () => {
    const service = buildService();
    const errorSpy = jest.spyOn(
      (service as unknown as { logger: { error: (...a: unknown[]) => void } }).logger,
      'error',
    );
    const warnSpy = jest.spyOn(
      (service as unknown as { logger: { warn: (...a: unknown[]) => void } }).logger,
      'warn',
    );

    (
      service as unknown as {
        logEmptyCompletion: (
          model: string,
          response: {
            choices?: Array<{ finishReason?: string }>;
            usage?: {
              completionTokens?: number;
              completionTokensDetails?: { reasoningTokens?: number };
            };
          },
          attempt: string,
        ) => void;
      }
    ).logEmptyCompletion(
      'openai/gpt-5',
      {
        choices: [{ finishReason: 'length' }],
        usage: {
          completionTokens: 960,
          completionTokensDetails: { reasoningTokens: 960 },
        },
      },
      'first attempt',
    );

    expect(warnSpy).toHaveBeenCalled();
    expect(errorSpy).toHaveBeenCalledWith(
      expect.stringContaining('ALERT reasoning_token_exhaustion'),
    );
  });
});
