import { AgenticLoopService } from '../src/agents/agenticLoop.service';
import { ExecutorAgent } from '../src/agents/executor.agent';
import { VerifierAgent } from '../src/agents/verifier.agent';
import { FormulaAnalyzer } from '../src/formula/formula.analyzer';
import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import { ToolBridgeService } from '../src/agents/tool-bridge.service';
import { CompletenessChecker } from '../src/agents/checkers/completeness.checker';
import { FormattingChecker } from '../src/agents/checkers/formatting.checker';
import { SseEmitter } from '../src/agents/sse.emitter';
import { Action, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

const baseContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:C3',
      rowCount: 3,
      columnCount: 3,
      values: [
        ['Name', 'Qty', 'Price'],
        ['Apple', 10, 1.5],
        ['Banana', 5, 0.75],
      ],
      formulas: [['', '', ''], ['', '', ''], ['', '', '']],
      numberFormats: [['General', 'General', 'General']],
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

const subtasks: SubTask[] = [
  {
    id: 's1',
    description: 'Add GST row',
    targetSheet: 'Sheet1',
    dependsOn: [],
    estimatedActions: 1,
  },
  {
    id: 's2',
    description: 'Format header',
    targetSheet: 'Sheet1',
    dependsOn: ['s1'],
    estimatedActions: 1,
  },
];

describe('AgenticLoopService verifier retry', () => {
  let executor: jest.Mocked<Pick<ExecutorAgent, 'execute'>>;
  let verifier: jest.Mocked<Pick<VerifierAgent, 'verify'>>;
  let formulaAnalyzer: jest.Mocked<Pick<FormulaAnalyzer, 'analyzeSheet'>>;
  let formulaValidator: jest.Mocked<
    Pick<
      FormulaValidatorService,
      'validatePreApply' | 'checkPostApply' | 'formatFeedback' | 'summarizeForVerifier'
    >
  >;
  let toolBridge: jest.Mocked<Pick<ToolBridgeService, 'waitForRangeData'>>;
  let completenessChecker: CompletenessChecker;
  let formattingChecker: FormattingChecker;
  let service: AgenticLoopService;
  let events: string[];

  beforeEach(() => {
    executor = { execute: jest.fn() };
    verifier = { verify: jest.fn() };
    formulaAnalyzer = { analyzeSheet: jest.fn().mockReturnValue({ llmSummary: '' }) };
    formulaValidator = {
      validatePreApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'pre_apply' }),
      checkPostApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'post_apply' }),
      formatFeedback: jest.fn().mockReturnValue(''),
      summarizeForVerifier: jest
        .fn()
        .mockReturnValue('Deterministic formula validator: all checks passed.'),
    };
    toolBridge = {
      waitForRangeData: jest.fn(),
    };
    completenessChecker = new CompletenessChecker();
    formattingChecker = new FormattingChecker();
    service = new AgenticLoopService(
      executor as unknown as ExecutorAgent,
      verifier as unknown as VerifierAgent,
      formulaAnalyzer as unknown as FormulaAnalyzer,
      formulaValidator as unknown as FormulaValidatorService,
      toolBridge as unknown as ToolBridgeService,
      completenessChecker,
      formattingChecker,
    );
    events = [];
  });

  const emit = (event: string, _data: Record<string, unknown>) => {
    events.push(event);
  };

  it('preserves passing subtask actions and only re-runs failing subtasks', async () => {
    const s1Action: Action = { type: 'ADD_ROW', data: ['GST', '', ''] } as Action;
    const s2ActionBad: Action = { type: 'SET_CELL', row: 99, col: 0, value: 'X' } as Action;
    const s2ActionFixed: Action = { type: 'SET_CELL', row: 0, col: 0, value: 'Header' } as Action;

    executor.execute
      .mockResolvedValueOnce({ subtaskId: 's1', actions: [s1Action], isDone: true })
      .mockResolvedValueOnce({ subtaskId: 's2', actions: [s2ActionBad], isDone: true })
      .mockResolvedValueOnce({ subtaskId: 's2', actions: [s2ActionFixed], isDone: true });

    verifier.verify
      .mockResolvedValueOnce({
        passed: false,
        feedback: 'Header row wrong',
        issues: [],
        subtaskResults: [
          { subtaskId: 's1', passed: true, feedback: 'OK', issues: [] },
          {
            subtaskId: 's2',
            passed: false,
            feedback: 'Wrong row',
            issues: [
              {
                severity: 'error',
                actionIndex: 0,
                description: 'Wrong row',
                suggestion: 'Use row 0',
              },
            ],
          },
        ],
      })
      .mockResolvedValueOnce({
        passed: true,
        feedback: 'All good',
        issues: [],
        subtaskResults: [
          { subtaskId: 's1', passed: true, feedback: 'OK', issues: [] },
          { subtaskId: 's2', passed: true, feedback: 'Fixed', issues: [] },
        ],
      });

    const result = await service.run(
      'Add GST and format header',
      subtasks,
      baseContext,
      new SseEmitter(emit),
    );

    expect(result.verifierPassed).toBe(true);
    expect(result.actions).toHaveLength(2);
    expect(result.actions[0]).toEqual(s1Action);
    expect(result.actions[1]).toEqual(s2ActionFixed);
    expect(executor.execute).toHaveBeenCalledTimes(3);
    expect(verifier.verify).toHaveBeenCalledTimes(2);
  });

  it('does not wipe all actions when LLM verifier fails', async () => {
    const action: Action = {
      type: 'SET_FORMULA',
      sheetName: 'Sheet1',
      row: 2,
      col: 2,
      formula: '=B2*C2',
    } as Action;

    executor.execute.mockResolvedValue({
      subtaskId: 's1',
      actions: [action],
      isDone: true,
    });

    verifier.verify.mockResolvedValue({
      passed: false,
      feedback: 'Minor issue',
      issues: [],
      subtaskResults: [
        {
          subtaskId: 's1',
          passed: false,
          feedback: 'Minor issue',
          issues: [],
        },
      ],
    });

    const singleSubtask = [subtasks[0]];
    const result = await service.run(
      'Add total formula',
      singleSubtask,
      baseContext,
      new SseEmitter(emit),
    );

    expect(result.actions).toEqual([action]);
    expect(result.verifierPassed).toBe(false);
    expect(executor.execute).toHaveBeenCalledTimes(2);
    expect(verifier.verify).toHaveBeenCalledTimes(2);
  });

  it('blocks invalid actions and retries executor with formula feedback', async () => {
    const validAction: Action = { type: 'SET_CELL', row: 0, col: 0, value: 'Header' } as Action;
    const badFormula: Action = {
      type: 'SET_FORMULA',
      sheetName: 'Sheet1',
      row: 1,
      col: 2,
      formula: '=Z99*2',
    } as Action;

    executor.execute
      .mockResolvedValueOnce({ subtaskId: 's1', actions: [badFormula], isDone: true })
      .mockResolvedValueOnce({ subtaskId: 's1', actions: [validAction], isDone: true });

    formulaValidator.validatePreApply
      .mockReturnValueOnce({
        passed: false,
        phase: 'pre_apply',
        issues: [
          {
            severity: 'error',
            code: 'REFERENCE',
            message: 'Out of bounds',
          },
        ],
      })
      .mockReturnValueOnce({ passed: true, issues: [], phase: 'pre_apply' });

    verifier.verify.mockResolvedValue({
      passed: true,
      feedback: 'OK',
      issues: [],
      subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'OK', issues: [] }],
    });

    const result = await service.run(
      'Fix header',
      [subtasks[0]],
      baseContext,
      new SseEmitter(emit),
    );

    expect(executor.execute).toHaveBeenCalledTimes(2);
    expect(result.actions).toEqual([validAction]);
  });

  it('fetches range data when executor requests get_range_data', async () => {
    const sortAction: Action = {
      type: 'SORT_RANGE',
      sheetName: 'Sheet1',
      range: 'A1:C100',
      key: 1,
      ascending: true,
      hasHeaders: true,
    } as Action;

    executor.execute
      .mockResolvedValueOnce({
        subtaskId: 's1',
        actions: [],
        isDone: false,
        toolRequest: { name: 'get_range_data', sheet: 'Sheet1', range: 'A1:C100' },
      })
      .mockResolvedValueOnce({ subtaskId: 's1', actions: [sortAction], isDone: true });

    toolBridge.waitForRangeData.mockResolvedValue({
      values: [
        ['Name', 'Qty', 'Price'],
        ['Apple', 10, 1.5],
      ],
    });

    verifier.verify.mockResolvedValue({
      passed: true,
      feedback: 'OK',
      issues: [],
      subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'OK', issues: [] }],
    });

    const result = await service.run(
      'Sort by qty',
      [subtasks[0]],
      { ...baseContext, onDemandFetchEnabled: true },
      new SseEmitter(emit),
      { conversationId: 'conv_test', toolEmit: emit },
    );

    expect(toolBridge.waitForRangeData).toHaveBeenCalledTimes(1);
    expect(result.actions).toEqual([sortAction]);
  });

  it('skips LLM verifier when deterministic checks pass for simple writes', async () => {
    const action: Action = { type: 'ADD_ROW', data: ['Total', 15, 18.75] } as Action;

    executor.execute.mockResolvedValue({
      subtaskId: 's1',
      actions: [action],
      isDone: true,
    });

    const result = await service.run(
      'Add total row',
      [subtasks[0]],
      baseContext,
      new SseEmitter(emit),
    );

    expect(result.verifierPassed).toBe(true);
    expect(verifier.verify).not.toHaveBeenCalled();
  });

  it('uses deterministic actions for create sheet + sort without calling the executor', async () => {
    const gstContext: WorkbookContext = {
      activeSheetName: 'Invoices',
      sheets: [
        {
          name: 'Invoices',
          usedRange: 'A1:M339',
          rowCount: 339,
          columnCount: 13,
          values: [
            ['Invoice No', 'Date', 'Customer', 'Amount', 'SGST', 'CGST', 'IGST'],
            ['INV-1', '2024-01-01', 'Customer', 1000, 90, 50, 0],
          ],
          formulas: [],
          numberFormats: [],
          structure: 'data_table',
        },
      ],
      namedRanges: [],
      tables: [],
    };

    const compoundSubtasks: SubTask[] = [
      {
        id: 's1',
        description: 'Create new sheet "CGST Sorted" as a copy of "Invoices"',
        targetSheet: 'Invoices',
        dependsOn: [],
        estimatedActions: 1,
      },
      {
        id: 's2',
        description: 'sort the values of CGST in ascending order on sheet "CGST Sorted"',
        targetSheet: 'CGST Sorted',
        dependsOn: ['s1'],
        estimatedActions: 1,
      },
    ];

    verifier.verify.mockResolvedValue({
      passed: true,
      feedback: 'OK',
      issues: [],
      subtaskResults: [
        { subtaskId: 's1', passed: true, feedback: 'OK', issues: [] },
        { subtaskId: 's2', passed: true, feedback: 'OK', issues: [] },
      ],
    });

    const result = await service.run(
      'create new sheet and sort the values of CGST in ascending order',
      compoundSubtasks,
      gstContext,
      new SseEmitter(emit),
    );

    expect(executor.execute).not.toHaveBeenCalled();
    expect(result.actions).toHaveLength(2);
    expect(result.actions[0].type).toBe('ADD_SHEET');
    expect(result.actions[1].type).toBe('SORT_RANGE');
    expect(result.actions[1].columnName).toBe('CGST');
    expect(result.verifierPassed).toBe(true);
  });
});
