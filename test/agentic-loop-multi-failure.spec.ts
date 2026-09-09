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
import { Action, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Live incident (TASKS.md #195): a stepwise wave of 12 independent
 * month-sheet-create subtasks (no dependsOn between them, so
 * `computeExecutionWaves` groups them into ONE parallel wave) had 9 of the
 * 12 genuinely fail — one after exhausting its retries, the other 8 never
 * even produced an action. Only 3 sheets (January, February, July) were
 * created; the other 9 months silently vanished with only ONE failure
 * reason ever recorded anywhere (`AgenticLoopResult.failedSubtask` is a
 * single field), because `buildLoopResult` used `.find()` to pick the first
 * failed subtask and discarded the rest.
 */

const baseContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:C3',
      rowCount: 3,
      columnCount: 3,
      values: [['Name', 'Qty', 'Price']],
      formulas: [['', '', '']],
      numberFormats: [['General', 'General', 'General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('AgenticLoopService — multiple independent subtasks failing in one wave', () => {
  let executor: jest.Mocked<Pick<ExecutorAgent, 'execute' | 'retryStep'>>;
  let verifier: jest.Mocked<Pick<VerifierAgent, 'verify'>>;
  let formulaAnalyzer: jest.Mocked<Pick<FormulaAnalyzer, 'analyzeSheet'>>;
  let formulaValidator: jest.Mocked<
    Pick<
      FormulaValidatorService,
      'validatePreApply' | 'checkPostApply' | 'formatFeedback' | 'summarizeForVerifier'
    >
  >;
  let toolBridge: jest.Mocked<Pick<ToolBridgeService, 'waitForRangeData'>>;
  let service: AgenticLoopService;

  beforeEach(() => {
    executor = { execute: jest.fn(), retryStep: jest.fn() };
    executor.retryStep.mockImplementation(async (retryContext, context, previousActions) =>
      executor.execute(retryContext.originalStep, context, previousActions),
    );
    verifier = { verify: jest.fn() };
    formulaAnalyzer = { analyzeSheet: jest.fn().mockReturnValue({ llmSummary: '' }) };
    formulaValidator = {
      validatePreApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'pre_apply' }),
      checkPostApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'post_apply' }),
      formatFeedback: jest.fn().mockReturnValue(''),
      summarizeForVerifier: jest.fn().mockReturnValue('ok'),
    };
    toolBridge = { waitForRangeData: jest.fn() };
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

  it('reports EVERY failed subtask, not just the first, when several independent subtasks in one wave never verify', async () => {
    // 12 independent month-sheet subtasks, exactly like the live incident.
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December',
    ];
    const subtasks: SubTask[] = months.map((month, i) => ({
      id: `s${i + 1}`,
      description: `Create ${month} sheet`,
      targetSheet: month,
      dependsOn: [],
      estimatedActions: 1,
    }));

    // 3 succeed (January=s1, February=s2, July=s7); the other 9 always come
    // back broken from the Executor, exhausting scoped retries.
    const succeedingIds = new Set(['s1', 's2', 's7']);
    executor.execute.mockImplementation(async (subtask: SubTask) => ({
      subtaskId: subtask.id,
      actions: [
        { type: 'ADD_SHEET', sheetName: subtask.targetSheet } as Action,
      ],
      isDone: true,
    }));

    // Every verify cycle reports the same 9 ids as failed — they never
    // recover, so scoped retry eventually exhausts MAX_STEP_RETRIES for all
    // of them in the same run.
    verifier.verify.mockImplementation(async (_prompt, verifySubtasks) => {
      const ids = (verifySubtasks as SubTask[]).map((s) => s.id);
      return {
        passed: ids.every((id) => succeedingIds.has(id)),
        feedback: 'Some months failed',
        issues: [],
        subtaskResults: ids.map((id) => ({
          subtaskId: id,
          passed: succeedingIds.has(id),
          feedback: succeedingIds.has(id) ? 'OK' : 'Sheet content invalid',
          issues: [],
        })),
      };
    });

    const result = await service.run(
      'Create 12 month sheets',
      subtasks,
      baseContext,
      new SseEmitter(emit),
    );

    expect(result.verifierPassed).toBe(false);

    // The core regression: EVERY failing subtask must have its own recorded
    // reason, not just one.
    const failedIds = result.failedSubtasks.map((f) => f.subtaskId).sort();
    const expectedFailedIds = months
      .map((_, i) => `s${i + 1}`)
      .filter((id) => !succeedingIds.has(id))
      .sort();
    expect(failedIds).toEqual(expectedFailedIds);
    expect(result.failedSubtasks.length).toBe(9);
    expect(result.failedSubtasks.every((f) => f.reason && f.reason.length > 0)).toBe(true);

    // Backward-compat single field still reports A failure (not null) —
    // existing single-failure callers keep working.
    expect(result.failedSubtask).not.toBeNull();
    expect(expectedFailedIds).toContain(result.failedSubtask!.subtaskId);

    // The 3 genuinely successful months still deliver their actions.
    const deliveredSheets = result.completedSubtasks.flatMap((c) =>
      c.actions.map((a) => (a as { sheetName?: string }).sheetName),
    );
    expect(deliveredSheets).toEqual(
      expect.arrayContaining(['January', 'February', 'July']),
    );
  });
});
