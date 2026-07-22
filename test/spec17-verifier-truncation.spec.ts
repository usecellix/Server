import { VerifierAgent } from '../src/agents/verifier.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { normalizeVerifierOutput } from '../src/agents/prompts/verifier.prompt';
import { salvageVerifierSubtaskResults } from '../src/agents/utils/verifier-partial-parse.util';
import { resolveVerifierMaxTokens } from '../src/agents/utils/verifier-token-budget.util';
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';
import { ExecutorAgent } from '../src/agents/executor.agent';
import { isExecutorBlockedSignal } from '../src/agents/utils/verifier-partial-parse.util';

const context: WorkbookContext = {
  activeSheetName: 'Dashboard',
  sheets: [
    {
      name: 'Dashboard',
      usedRange: 'A1:B2',
      rowCount: 2,
      columnCount: 2,
      values: [
        ['A', 'B'],
        [1, 2],
      ],
      formulas: [
        ['', ''],
        ['', ''],
      ],
      numberFormats: [
        ['General', 'General'],
        ['General', 'General'],
      ],
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

const twelveSubtasks: SubTask[] = Array.from({ length: 12 }, (_, i) => ({
  id: `s${i + 1}`,
  description: `Step s${i + 1}`,
  targetSheet: 'Dashboard',
  dependsOn: [],
  estimatedActions: 1,
}));

/** Spec 17 truncated fixture: s1–s8 complete, cut off mid-s9 feedback. */
const TRUNCATED_VERIFIER_RAW = `{
  "passed": false,
  "feedback": "s9/s10 issues",
  "issues": [],
  "subtaskResults": [
    {"subtaskId":"s1","passed":true,"feedback":"ok","issues":[]},
    {"subtaskId":"s2","passed":true,"feedback":"ok","issues":[]},
    {"subtaskId":"s3","passed":true,"feedback":"ok","issues":[]},
    {"subtaskId":"s4","passed":true,"feedback":"ok","issues":[]},
    {"subtaskId":"s5","passed":true,"feedback":"ok","issues":[]},
    {"subtaskId":"s6","passed":true,"feedback":"ok","issues":[]},
    {"subtaskId":"s7","passed":true,"feedback":"ok","issues":[]},
    {"subtaskId":"s8","passed":true,"feedback":"ok","issues":[]},
    {"subtaskId":"s9","passed":false,"feedback":`;

describe('Spec 17 verifier truncation', () => {
  it('scales max tokens with subtask count', () => {
    expect(resolveVerifierMaxTokens(1)).toBeGreaterThanOrEqual(2000);
    expect(resolveVerifierMaxTokens(12)).toBeGreaterThan(resolveVerifierMaxTokens(1));
    expect(resolveVerifierMaxTokens(12)).toBeGreaterThanOrEqual(3000);
  });

  it('salvages complete subtaskResults from truncated JSON', () => {
    const salvaged = salvageVerifierSubtaskResults(TRUNCATED_VERIFIER_RAW);
    expect(salvaged.map((r) => r.subtaskId)).toEqual([
      's1',
      's2',
      's3',
      's4',
      's5',
      's6',
      's7',
      's8',
    ]);
    expect(salvaged.every((r) => r.passed)).toBe(true);
  });

  it('marks missing IDs inconclusive instead of blanket-failing all', () => {
    const salvaged = salvageVerifierSubtaskResults(TRUNCATED_VERIFIER_RAW);
    const normalized = normalizeVerifierOutput(
      {
        passed: false,
        feedback: 'truncated',
        issues: [],
        subtaskResults: salvaged,
      },
      twelveSubtasks.map((s) => s.id),
      { fillMissingAsInconclusive: true },
    );

    for (const id of ['s1', 's2', 's3', 's4', 's5', 's6', 's7', 's8']) {
      const row = normalized.subtaskResults.find((r) => r.subtaskId === id);
      expect(row?.passed).toBe(true);
      expect(row?.inconclusive).toBeFalsy();
    }
    for (const id of ['s9', 's10', 's11', 's12']) {
      const row = normalized.subtaskResults.find((r) => r.subtaskId === id);
      expect(row?.inconclusive).toBe(true);
      expect(row?.passed).toBe(false);
    }
    expect(normalized.passed).toBe(false);
  });

  it('retries verify-only with higher budget then preserves salvaged passes', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce(TRUNCATED_VERIFIER_RAW)
      .mockResolvedValueOnce(
        JSON.stringify({
          passed: false,
          feedback: 's9 s10 still fail',
          issues: [],
          subtaskResults: twelveSubtasks.map((s) => ({
            subtaskId: s.id,
            passed: s.id !== 's9' && s.id !== 's10',
            feedback: s.id === 's9' || s.id === 's10' ? 'fail' : 'ok',
            issues: [],
          })),
        }),
      );

    const agent = new VerifierAgent(
      { complete } as unknown as OpenRouterService,
      { openRouterModelMedium: 'openai/gpt-5-mini' } as unknown as AppConfigService,
    );

    const actionsBySubtask = Object.fromEntries(
      twelveSubtasks.map((s) => [s.id, [{ type: 'SET_CELL', value: s.id }]]),
    );

    const result = await agent.verify(
      'dashboard chart analysis',
      twelveSubtasks,
      actionsBySubtask as never,
      context,
      undefined,
      'corr_v1',
    );

    expect(complete).toHaveBeenCalledTimes(2);
    expect(complete.mock.calls[0][0].maxTokens).toBeGreaterThanOrEqual(3000);
    expect(complete.mock.calls[1][0].maxTokens).toBe(8192);

    for (const id of ['s1', 's2', 's3', 's4', 's5', 's6', 's7', 's8', 's11', 's12']) {
      expect(result.subtaskResults.find((r) => r.subtaskId === id)?.passed).toBe(true);
    }
    expect(result.subtaskResults.find((r) => r.subtaskId === 's9')?.passed).toBe(false);
    expect(result.subtaskResults.find((r) => r.subtaskId === 's10')?.passed).toBe(false);
  });
});

describe('Spec 17 Bug C — blocked signal not overridden', () => {
  it('detects blocked nextStep', () => {
    expect(
      isExecutorBlockedSignal(
        "Blocked: AGGREGATE_TABLE cannot group by MONTH(Date) without a helper column, and on-demand fetch is disabled... Add a 'Month' column... or enable data fetch.",
      ),
    ).toBe(true);
    expect(isExecutorBlockedSignal('Continue with next rows')).toBe(false);
  });

  it('does not apply SORT_RANGE fallback when executor reports Blocked', async () => {
    const blockedNext =
      "Blocked: AGGREGATE_TABLE cannot group by MONTH(Date) without a helper column, and on-demand fetch is disabled... Add a 'Month' column... or enable data fetch.";

    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({
        subtaskId: 's9',
        actions: [],
        isDone: false,
        nextStep: blockedNext,
      }),
    );

    const agent = new ExecutorAgent(
      { complete } as unknown as OpenRouterService,
      { openRouterModelHigh: 'openai/gpt-5' } as unknown as AppConfigService,
    );

    const result = await agent.execute(
      {
        id: 's9',
        description: 'sort the values of Dashboard by amount',
        targetSheet: 'Dashboard',
        dependsOn: [],
        estimatedActions: 1,
      },
      context,
      [],
      'corr_block',
    );

    expect(result.isDone).toBe(false);
    expect(result.actions).toHaveLength(0);
    expect(result.nextStep).toBe(blockedNext);
    expect(result.actions.some((a) => a.type === 'SORT_RANGE')).toBe(false);
  });
});
