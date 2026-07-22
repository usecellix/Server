import { Tier2GenerateVerifyService } from '../src/excel-ai/services/tier2-generate-verify.service';
import { ExecutorAgent } from '../src/agents/executor.agent';
import { VerifierAgent } from '../src/agents/verifier.agent';
import { ToolBridgeService } from '../src/agents/tool-bridge.service';
import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import { WorkbookContext } from '../src/agents/types/agent.types';
import {
  collectRecentTurnActionRecords,
  extractTurnActionRecords,
  formatTurnActionRecordsForExecutor,
  referencesPriorChartOrTable,
} from '../src/excel-ai/utils/turn-action-history.util';

const workbookContext: WorkbookContext = {
  activeSheetName: 'Dashboard',
  sheets: [
    {
      name: 'Dashboard',
      usedRange: 'A1:B10',
      rowCount: 10,
      columnCount: 2,
      values: [
        ['Month', 'Total'],
        ['Jan', 10],
        ['Feb', 20],
      ],
      formulas: [],
      numberFormats: [],
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('Tier2GenerateVerifyService Spec 18 retry', () => {
  function buildService(mocks: {
    execute: jest.Mock;
    retryStep: jest.Mock;
    verify: jest.Mock;
    waitForRangeData?: jest.Mock;
  }): Tier2GenerateVerifyService {
    const executor = {
      execute: mocks.execute,
      retryStep: mocks.retryStep,
      modelName: 'openai/gpt-5',
    } as unknown as ExecutorAgent;
    const verifier = { verify: mocks.verify } as unknown as VerifierAgent;
    const formulaValidator = {
      checkNoHardcodedLiterals: jest.fn().mockReturnValue({ passed: true }),
      validatePreApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'pre_apply' }),
      summarizeForVerifier: jest.fn().mockReturnValue('ok'),
      formatFeedback: jest.fn().mockReturnValue(''),
    } as unknown as FormulaValidatorService;
    const toolBridge = {
      waitForRangeData:
        mocks.waitForRangeData ??
        jest.fn().mockResolvedValue({ values: [['Month', 'Total'], ['Jan', 10]] }),
    } as unknown as ToolBridgeService;
    return new Tier2GenerateVerifyService(executor, verifier, formulaValidator, toolBridge);
  }

  it('retries exactly once with verifier suggestion then passes (wrong sourceRange fixture)', async () => {
    const badAction = {
      type: 'CREATE_CHART',
      sheetName: 'Dashboard',
      sourceSheetName: 'Dashboard',
      sourceRange: 'A4:B54',
      chartType: 'BarClustered',
      chartId: 'Chart_bad',
      destCell: 'D2',
    };
    const goodAction = {
      ...badAction,
      sourceRange: 'A1:B10',
      chartId: 'Chart_fixed',
    };

    const execute = jest.fn().mockResolvedValue({
      subtaskId: 's1',
      actions: [badAction],
      isDone: true,
    });
    const retryStep = jest.fn().mockResolvedValue({
      subtaskId: 's1',
      actions: [goodAction],
      isDone: true,
    });
    const verify = jest
      .fn()
      .mockResolvedValueOnce({
        passed: false,
        feedback:
          "Chart action will be created but the sourceRange likely omits existing data (starts at A4).",
        issues: [
          {
            severity: 'error',
            description:
              "sourceRange 'A4:B54' begins at row 4 and will omit the header and data in rows 2-3",
            suggestion: "Use a sourceRange that includes the header and all data, e.g. 'A1:B10'",
          },
        ],
        subtaskResults: [
          {
            subtaskId: 's1',
            passed: false,
            feedback: 'bad range',
            issues: [
              {
                severity: 'error',
                description: "sourceRange 'A4:B54' begins at row 4",
                suggestion: "Use a sourceRange that includes the header and all data, e.g. 'A1:B10'",
              },
            ],
          },
        ],
      })
      .mockResolvedValueOnce({
        passed: true,
        feedback: 'ok',
        issues: [],
        subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'ok', issues: [] }],
      });

    const service = buildService({ execute, retryStep, verify });
    const result = await service.execute(
      'Also create a bar chart along with the current',
      'CREATE_CHART',
      workbookContext,
      'corr_t2',
    );

    expect(result.verifierPassed).toBe(true);
    expect(result.retried).toBe(true);
    expect(result.actions[0]).toMatchObject({ sourceRange: 'A1:B10' });
    expect(execute).toHaveBeenCalledTimes(1);
    expect(retryStep).toHaveBeenCalledTimes(1);
    expect(verify).toHaveBeenCalledTimes(2);

    const retryCtx = retryStep.mock.calls[0][1] as WorkbookContext;
    expect(retryCtx.verifierFeedback).toContain('A1:B10');
    expect(retryCtx.verifierFeedback).toContain('Suggestion:');
  });

  it('surfaces verifier suggestion in failure answer after single retry still fails', async () => {
    const badAction = {
      type: 'CREATE_CHART',
      sheetName: 'Dashboard',
      sourceRange: 'A4:B54',
      chartType: 'BarClustered',
    };
    const execute = jest.fn().mockResolvedValue({
      subtaskId: 's1',
      actions: [badAction],
      isDone: true,
    });
    const retryStep = jest.fn().mockResolvedValue({
      subtaskId: 's1',
      actions: [badAction],
      isDone: true,
    });
    const failResult = {
      passed: false,
      feedback: 'sourceRange still wrong',
      issues: [
        {
          severity: 'error' as const,
          description: 'range starts at A4',
          suggestion: "try A1:B10",
        },
      ],
      subtaskResults: [
        {
          subtaskId: 's1',
          passed: false,
          feedback: 'still wrong',
          issues: [
            {
              severity: 'error' as const,
              description: 'range starts at A4',
              suggestion: 'try A1:B10',
            },
          ],
        },
      ],
    };
    const verify = jest.fn().mockResolvedValue(failResult);

    const service = buildService({ execute, retryStep, verify });
    const result = await service.execute('make a bar chart', 'CREATE_CHART', workbookContext);

    expect(result.verifierPassed).toBe(false);
    expect(result.retried).toBe(true);
    expect(result.answer).toContain('try A1:B10');
    expect(verify).toHaveBeenCalledTimes(2);
  });

  it('Bug 4: retry toolRequest → resolve get_range_data → one follow-up Executor → pass (green bar graph)', async () => {
    const badCreate = {
      type: 'CREATE_CHART',
      sheetName: 'Dashboard',
      sourceSheetName: 'Dashboard',
      sourceRange: 'A4:B54',
      chartType: 'BarClustered',
      chartId: 'Chart_bad',
      destCell: 'D2',
    };
    const greenUpdate = {
      type: 'UPDATE_CHART',
      sheetName: 'Dashboard',
      chartId: 'Chart_bad',
      chartType: 'BarClustered',
      colorScheme: 'green',
    };
    const fixedCreate = {
      ...badCreate,
      sourceRange: 'A1:B10',
      chartId: 'Chart_green_bar',
      colorScheme: 'green',
    };
    const fixedUpdate = {
      ...greenUpdate,
      chartId: 'Chart_green_bar',
      colorScheme: 'green',
    };

    const execute = jest
      .fn()
      .mockResolvedValueOnce({
        subtaskId: 's1',
        actions: [badCreate, greenUpdate],
        isDone: true,
      })
      // Bug 4 follow-up after tool data
      .mockResolvedValueOnce({
        subtaskId: 's1',
        actions: [fixedCreate, fixedUpdate],
        isDone: true,
      });
    const retryStep = jest.fn().mockResolvedValue({
      subtaskId: 's1',
      actions: [],
      isDone: false,
      toolRequest: { name: 'get_range_data', sheet: 'Dashboard', range: 'A1:B60' },
    });
    const waitForRangeData = jest.fn().mockResolvedValue({
      values: [
        ['Month', 'Total'],
        ['Jan', 10],
        ['Feb', 20],
      ],
    });
    const verify = jest
      .fn()
      .mockResolvedValueOnce({
        passed: false,
        feedback: 'sourceRange omits header',
        issues: [
          {
            severity: 'error',
            description: "sourceRange 'A4:B54' begins at row 4",
            suggestion: "Use 'A1:B10'",
          },
        ],
        subtaskResults: [
          {
            subtaskId: 's1',
            passed: false,
            feedback: 'bad range',
            issues: [
              {
                severity: 'error',
                description: "sourceRange 'A4:B54' begins at row 4",
                suggestion: "Use 'A1:B10'",
              },
            ],
          },
        ],
      })
      .mockResolvedValueOnce({
        passed: true,
        feedback: 'ok',
        issues: [],
        subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'ok', issues: [] }],
      });

    const service = buildService({ execute, retryStep, verify, waitForRangeData });
    const emit = jest.fn();
    const result = await service.execute(
      'Create a bar graph too using color green',
      'CREATE_CHART',
      workbookContext,
      'corr_bug4',
      { conversationId: 'conv_bug4', toolEmit: emit },
    );

    expect(result.verifierPassed).toBe(true);
    expect(result.retried).toBe(true);
    expect(result.toolFollowUp).toBe(true);
    expect(result.actions.length).toBeGreaterThan(0);
    expect(result.actions).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ type: 'CREATE_CHART', sourceRange: 'A1:B10' }),
        expect.objectContaining({ type: 'UPDATE_CHART', colorScheme: 'green' }),
      ]),
    );
    expect(retryStep).toHaveBeenCalledTimes(1);
    expect(waitForRangeData).toHaveBeenCalledTimes(1);
    expect(waitForRangeData.mock.calls[0][1]).toMatchObject({
      name: 'get_range_data',
      sheet: 'Dashboard',
      range: 'A1:B60',
    });
    // Initial execute + one tool-informed follow-up only (cap = 2 correction calls: retry + follow-up)
    expect(execute).toHaveBeenCalledTimes(2);
    expect(verify).toHaveBeenCalledTimes(2);

    const followUpCtx = execute.mock.calls[1][1] as WorkbookContext;
    expect(followUpCtx.fetchedRanges).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ sheet: 'Dashboard', range: 'A1:B60' }),
      ]),
    );
  });

  it('Bug 4: does not open-end loop — at most one tool-informed follow-up', async () => {
    const badAction = {
      type: 'CREATE_CHART',
      sheetName: 'Dashboard',
      sourceRange: 'A4:B54',
      chartType: 'BarClustered',
    };
    const execute = jest
      .fn()
      .mockResolvedValueOnce({
        subtaskId: 's1',
        actions: [badAction],
        isDone: true,
      })
      // Follow-up still asks for another tool — must NOT recurse
      .mockResolvedValueOnce({
        subtaskId: 's1',
        actions: [],
        isDone: false,
        toolRequest: { name: 'get_range_data', sheet: 'Dashboard', range: 'A1:C100' },
      });
    const retryStep = jest.fn().mockResolvedValue({
      subtaskId: 's1',
      actions: [],
      isDone: false,
      toolRequest: { name: 'get_range_data', sheet: 'Dashboard', range: 'A1:B60' },
    });
    const waitForRangeData = jest.fn().mockResolvedValue({
      values: [['Month', 'Total'], ['Jan', 10]],
    });
    const verify = jest.fn().mockResolvedValue({
      passed: false,
      feedback: 'no actions were provided',
      issues: [],
      subtaskResults: [{ subtaskId: 's1', passed: false, feedback: 'no actions', issues: [] }],
    });

    const service = buildService({ execute, retryStep, verify, waitForRangeData });
    const result = await service.execute(
      'Create a bar graph too using color green',
      'CREATE_CHART',
      workbookContext,
      'corr_bug4_cap',
      { conversationId: 'conv_cap', toolEmit: jest.fn() },
    );

    expect(result.toolFollowUp).toBe(true);
    expect(result.verifierPassed).toBe(false);
    expect(waitForRangeData).toHaveBeenCalledTimes(1);
    expect(retryStep).toHaveBeenCalledTimes(1);
    expect(execute).toHaveBeenCalledTimes(2);
  });

  it('live repro: color retry then range fail → deterministic sourceRange patch passes', async () => {
    const firstActions = [
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Dashboard',
        sourceRange: 'A4:B54',
        chartType: 'BarClustered',
        chartId: 'Chart_GreenBar',
      },
    ];
    const retryActions = [
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Dashboard',
        sourceRange: 'A4:B54',
        chartType: 'BarClustered',
        chartId: 'Chart_GreenBar',
      },
      {
        type: 'UPDATE_CHART',
        sheetName: 'Dashboard',
        chartId: 'Chart_GreenBar',
        chartType: 'BarClustered',
        colorScheme: 'green',
      },
    ];

    const execute = jest.fn().mockResolvedValue({
      subtaskId: 's1',
      actions: firstActions,
      isDone: true,
    });
    const retryStep = jest.fn().mockResolvedValue({
      subtaskId: 's1',
      actions: retryActions,
      isDone: true,
    });
    const verify = jest
      .fn()
      .mockResolvedValueOnce({
        passed: false,
        feedback: 'CREATE_CHART does not specify color green',
        issues: [
          {
            severity: 'error',
            description: 'CREATE_CHART does not specify the chart/series color',
            suggestion: "set series fill to 'green'",
          },
        ],
        subtaskResults: [
          {
            subtaskId: 's1',
            passed: false,
            feedback: 'missing color',
            issues: [
              {
                severity: 'error',
                description: 'missing color',
                suggestion: "set fill to green",
              },
            ],
          },
        ],
      })
      .mockResolvedValueOnce({
        passed: false,
        feedback:
          'Chart creation action will likely miss data because the sourceRange starts at A4 while the workbook data begins in earlier rows. Color update is fine.',
        issues: [
          {
            severity: 'error',
            description: "CREATE_CHART sourceRange 'A4:B54' likely excludes existing data",
            suggestion:
              "Adjust sourceRange to include the actual data rows (for example 'A1:B10' or 'A2:B10')",
          },
        ],
        subtaskResults: [
          {
            subtaskId: 's1',
            passed: false,
            feedback: 'bad range',
            issues: [
              {
                severity: 'error',
                description: 'CREATE_CHART uses an incorrect sourceRange',
                suggestion: "Change sourceRange to cover the actual data (e.g., 'A1:B10')",
              },
            ],
          },
        ],
      })
      .mockResolvedValueOnce({
        passed: true,
        feedback: 'ok',
        issues: [],
        subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'ok', issues: [] }],
      });

    const service = buildService({ execute, retryStep, verify });
    const result = await service.execute(
      'Create a bar graph too using color green',
      'CREATE_CHART',
      workbookContext,
      'corr_live',
    );

    expect(result.verifierPassed).toBe(true);
    expect(result.deterministicPatch).toBe(true);
    expect(result.actions).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          type: 'CREATE_CHART',
          sourceRange: 'A1:B10',
        }),
        expect.objectContaining({
          colorScheme: 'green',
        }),
      ]),
    );
    // Initial verify + post-retry verify + post-patch verify
    expect(verify).toHaveBeenCalledTimes(3);
  });
});

describe('turn-action-history Spec 18 Bug 2', () => {
  it('extracts CREATE_CHART records with sourceRange and chartId', () => {
    const records = extractTurnActionRecords([
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Dashboard',
        sourceRange: 'A1:B10',
        chartId: 'Chart_top',
        chartType: 'ColumnClustered',
        destCell: 'D2',
      },
    ]);
    expect(records).toEqual([
      expect.objectContaining({
        actionType: 'CREATE_CHART',
        sourceRange: 'A1:B10',
        chartId: 'Chart_top',
        sheetName: 'Dashboard',
      }),
    ]);
  });

  it('collects records from assistant message metadata for follow-up turns', () => {
    const history = [
      {
        role: 'assistant',
        metadata: {
          actions: [
            {
              type: 'CREATE_CHART',
              sheetName: 'Dashboard',
              sourceRange: 'A1:B10',
              chartId: 'Chart_1',
            },
          ],
          turnActionRecords: [
            {
              actionType: 'CREATE_CHART' as const,
              sheetName: 'Dashboard',
              sourceRange: 'A1:B10',
              chartId: 'Chart_1',
            },
          ],
        },
      },
    ];
    const collected = collectRecentTurnActionRecords(history);
    expect(collected[0]?.sourceRange).toBe('A1:B10');
    expect(referencesPriorChartOrTable('Also create a bar chart along with the current')).toBe(
      true,
    );
    expect(formatTurnActionRecordsForExecutor(collected)).toContain('sourceRange=A1:B10');
    expect(formatTurnActionRecordsForExecutor(collected)).toContain('authoritative');
  });
});
