import { AgenticLoopService } from '../src/agents/agenticLoop.service';
import { ExecutorAgent } from '../src/agents/executor.agent';
import { VerifierAgent } from '../src/agents/verifier.agent';
import { FormulaAnalyzer } from '../src/formula/formula.analyzer';
import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import { ToolBridgeService } from '../src/agents/tool-bridge.service';
import { CompletenessChecker } from '../src/agents/checkers/completeness.checker';
import { FormattingChecker } from '../src/agents/checkers/formatting.checker';
import { OverwriteOccupancyChecker } from '../src/agents/checkers/overwrite-occupancy.checker';
import { SseEmitter } from '../src/agents/sse.emitter';
import { splitSpecPinnedSubtasks } from '../src/agents/utils/header-table-split.util';
import { Action, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md — two consecutive LIVE 12-month runs had the
 * (single, solo) template subtask hit "max iterations (10)" or time out, even
 * with no concurrency contention at all: create-sheet + header row + table +
 * formulas + dropdowns + widths + font, all in one Executor generation, was
 * too much work. This test proves the fix end to end through the real loop:
 * split into a zero-LLM-call header/table step plus a lighter "rest" step,
 * the wave now finishes even though the Executor would still fail if asked to
 * do everything at once.
 */

const baseContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [{
    name: 'Sheet1', usedRange: 'A1:A1', rowCount: 1, columnCount: 1,
    values: [['']], formulas: [['']], numberFormats: [['General']],
    structure: 'data_table', headerRowIndex: 0,
  }],
  namedRanges: [],
  tables: [],
};

describe('AgenticLoopService — deterministic header/table split (TASKS.md #280)', () => {
  let executor: jest.Mocked<Pick<ExecutorAgent, 'execute' | 'retryStep'>>;
  let verifier: jest.Mocked<Pick<VerifierAgent, 'verify'>>;
  let service: AgenticLoopService;

  beforeEach(() => {
    executor = { execute: jest.fn(), retryStep: jest.fn() };
    executor.retryStep.mockImplementation(async (retryContext, context, previousActions) =>
      executor.execute(retryContext.originalStep, context, previousActions),
    );
    verifier = { verify: jest.fn() };
    const formulaAnalyzer = { analyzeSheet: jest.fn().mockReturnValue({ llmSummary: '' }) };
    const formulaValidator = {
      validatePreApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'pre_apply' }),
      checkPostApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'post_apply' }),
      formatFeedback: jest.fn().mockReturnValue(''),
      summarizeForVerifier: jest.fn().mockReturnValue('ok'),
    };
    const toolBridge = { waitForRangeData: jest.fn() };
    service = new AgenticLoopService(
      executor as unknown as ExecutorAgent,
      verifier as unknown as VerifierAgent,
      formulaAnalyzer as unknown as FormulaAnalyzer,
      formulaValidator as unknown as FormulaValidatorService,
      toolBridge as unknown as ToolBridgeService,
      new CompletenessChecker(),
      new FormattingChecker(),
      new OverwriteOccupancyChecker(),
    );
  });

  const emit = () => {};

  const monolithic = (id: string, month: string): SubTask => ({
    id,
    description: `Create sheet '${month}' and build the payment ledger in one go: headers, table, formulas, dropdowns, widths, font.`,
    targetSheet: month,
    dependsOn: [],
    estimatedActions: 26,
    expectedHeaders: ['Unit No', 'Guest', 'Check In', 'Check Out', 'Total Amount'],
  });

  it('the header/table step never calls the Executor and completes at zero cost', async () => {
    const [headerStep] = splitSpecPinnedSubtasks([monolithic('p2_s1', 'January')]);

    const result = await service.run('Build January', [headerStep], baseContext, new SseEmitter(emit));

    expect(executor.execute).not.toHaveBeenCalled();
    expect(result.failedSubtasks).toEqual([]);
    const sheetAction = result.actions.find((a) => a.type === 'ADD_SHEET') as { name?: string } | undefined;
    expect(sheetAction?.name).toBe('January');
    const tableAction = result.actions.find((a) => a.type === 'CREATE_TABLE') as { tableName?: string } | undefined;
    expect(tableAction?.tableName).toBe('tblJanuary');
  });

  it('a subtask so heavy the Executor cannot finish it STILL succeeds once split (the actual regression)', async () => {
    // Reproduces the live shape exactly: given the FULL monolithic work in one
    // generation, the Executor never converges — this is deliberately the
    // same failure the two live runs hit. What changes is that this subtask
    // no longer has to do create+headers+table at all.
    const [headerStep, restStep] = splitSpecPinnedSubtasks([monolithic('p2_s1', 'January')]);

    executor.execute.mockImplementation(async (subtask: SubTask) => {
      // The rest step's LIGHTER job (formulas/dropdowns/widths/font only)
      // finishes in one shot — proving the split, not luck, is what fixed it.
      if (subtask.id === restStep.id) {
        return {
          subtaskId: subtask.id,
          actions: [{ type: 'SET_FORMULA', sheetName: 'January', row: 1, col: 5, formula: '=D2-C2' } as Action],
          isDone: true,
        };
      }
      return { subtaskId: subtask.id, actions: [], isDone: false, nextStep: 'still working' };
    });
    verifier.verify.mockResolvedValue({
      passed: true, feedback: 'ok', issues: [],
      subtaskResults: [{ subtaskId: restStep.id, passed: true, feedback: 'OK', issues: [] }],
    });

    // Wave 1: header step alone (as the real plan-wide wave computation would
    // schedule it — restStep depends on headerStep, one wave later).
    const wave1 = await service.run('Build January', [headerStep], baseContext, new SseEmitter(emit));
    expect(wave1.failedSubtasks).toEqual([]);

    // Wave 2: rest step, with wave 1's sheet/headers/table now "prior work" —
    // the same shape `runStepwiseWave` uses for every wave after the first.
    const wave2 = await service.runWave(
      'Build January',
      [restStep],
      wave1.completedSubtasks.map((c) => ({ subtask: headerStep, actions: c.actions })),
      baseContext,
      new SseEmitter(emit),
    );

    expect(wave2.failedSubtasks).toEqual([]);
    expect(wave2.actions.some((a) => a.type === 'SET_FORMULA')).toBe(true);
  });
});

/**
 * TASKS.md #302 — a subtask refused for having already done its own work.
 *
 * Live shape, from the third smoke run: ALL TWELVE month "rest" steps failed
 * to converge, and the refusal quoted the step's own formula back at it as the
 * existing value:
 *
 *   Write blocked: target range F2 already contains data. This action would
 *   overwrite existing values. Existing values include: =IF(OR(D2="",E2="")…
 *
 * `runDeterministicChecks` graded the occupancy checker against
 * `verifyContext` — the shadow enriched from EVERY state including the
 * subtask's own actions — so the subtask was graded against a workbook that
 * already contained the very writes being graded. Every retry was refused the
 * same way, all twelve gated off Main, and the run ended at wave 4 of 6.
 *
 * The check fires only for "add/insert column" descriptions, which is exactly
 * what a split rest step's description looks like ("Add only what is still
 * needed…", "…already has its header cell and column position").
 */
describe('AgenticLoopService — a retry is not an overwrite of itself (TASKS.md #302)', () => {
  let executor: jest.Mocked<Pick<ExecutorAgent, 'execute' | 'retryStep'>>;
  let verifier: jest.Mocked<Pick<VerifierAgent, 'verify'>>;
  let service: AgenticLoopService;

  beforeEach(() => {
    executor = { execute: jest.fn(), retryStep: jest.fn() };
    executor.retryStep.mockImplementation(async (retryContext, context, previousActions) =>
      executor.execute(retryContext.originalStep, context, previousActions),
    );
    verifier = { verify: jest.fn() };
    const formulaAnalyzer = { analyzeSheet: jest.fn().mockReturnValue({ llmSummary: '' }) };
    const formulaValidator = {
      validatePreApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'pre_apply' }),
      checkPostApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'post_apply' }),
      formatFeedback: jest.fn().mockReturnValue(''),
      summarizeForVerifier: jest.fn().mockReturnValue('ok'),
    };
    const toolBridge = { waitForRangeData: jest.fn() };
    service = new AgenticLoopService(
      executor as unknown as ExecutorAgent,
      verifier as unknown as VerifierAgent,
      formulaAnalyzer as unknown as FormulaAnalyzer,
      formulaValidator as unknown as FormulaValidatorService,
      toolBridge as unknown as ToolBridgeService,
      new CompletenessChecker(),
      new FormattingChecker(),
      new OverwriteOccupancyChecker(),
    );
  });

  const emit = () => {};

  /** A sheet that already has a header row and a seeded, EMPTY data row. */
  const contextWithSeededSheet: WorkbookContext = {
    ...baseContext,
    sheets: [
      ...baseContext.sheets,
      {
        name: 'January',
        usedRange: 'A1:F2',
        rowCount: 2,
        columnCount: 6,
        values: [
          ['Unit No', 'Guest', 'Check In', 'Check Out', 'Rate', 'Total Amount'],
          ['', '', '', '', '', ''],
        ],
        formulas: [
          ['', '', '', '', '', ''],
          ['', '', '', '', '', ''],
        ],
        numberFormats: [['General'], ['General']],
        structure: 'data_table',
        headerRowIndex: 0,
      },
    ],
  };

  /** The live rest-step description, which is what makes the checker fire. */
  const restStep: SubTask = {
    id: 'p2_s1',
    targetSheet: 'January',
    dependsOn: [],
    estimatedActions: 20,
    description:
      "Sheet 'January', its header row and table 'tblJanuary' already exist — do not rewrite the " +
      'header row: every column below, including any computed ones, already has its header cell ' +
      'and column position. Add only what is still needed: row-2 formulas, dropdowns, and ' +
      'set column widths A=80, B=80.',
  };

  it('the rest step completes instead of being refused for its own write', async () => {
    executor.execute.mockResolvedValue({
      subtaskId: restStep.id,
      actions: [
        {
          type: 'SET_FORMULA',
          sheetName: 'January',
          row: 1,
          col: 5,
          formula: '=IF(OR(C2="",D2=""),"",D2-C2)',
        } as Action,
      ],
      isDone: true,
    });
    verifier.verify.mockResolvedValue({
      passed: true,
      feedback: 'ok',
      issues: [],
      subtaskResults: [{ subtaskId: restStep.id, passed: true, feedback: 'OK', issues: [] }],
    });

    const result = await service.run(
      'Build January',
      [restStep],
      contextWithSeededSheet,
      new SseEmitter(emit),
    );

    expect(result.failedSubtasks).toEqual([]);
    expect(result.actions.some((a) => a.type === 'SET_FORMULA')).toBe(true);
  });

  it('a write onto a cell that was ALREADY occupied is still refused', async () => {
    // The check must keep doing its job: F2 holds real pre-existing data that
    // this run did not write, so overwriting it is the genuine mistake.
    const occupied: WorkbookContext = {
      ...contextWithSeededSheet,
      sheets: contextWithSeededSheet.sheets.map((sheet) =>
        sheet.name === 'January'
          ? {
              ...sheet,
              values: [
                ['Unit No', 'Guest', 'Check In', 'Check Out', 'Rate', 'Total Amount'],
                ['', '', '', '', '', 1200],
              ],
            }
          : sheet,
      ),
    };

    executor.execute.mockResolvedValue({
      subtaskId: restStep.id,
      actions: [
        { type: 'SET_FORMULA', sheetName: 'January', row: 1, col: 5, formula: '=D2-C2' } as Action,
      ],
      isDone: true,
    });
    verifier.verify.mockResolvedValue({
      passed: false,
      feedback: 'overwrite',
      issues: [],
      subtaskResults: [{ subtaskId: restStep.id, passed: false, feedback: 'overwrite', issues: [] }],
    });

    const result = await service.run('Build January', [restStep], occupied, new SseEmitter(emit));

    expect(result.failedSubtasks.length).toBeGreaterThan(0);
  });
});
