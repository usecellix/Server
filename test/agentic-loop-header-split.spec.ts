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
