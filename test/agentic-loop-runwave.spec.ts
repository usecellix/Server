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
 * STEPWISE_EXECUTION.md SD-1/SD-3 — `runWave` executes ONE dependency wave and
 * reports only that wave's work, while still seeing earlier accepted waves'
 * results in its shadow workbook and Executor context.
 */

const baseContext: WorkbookContext = {
  activeSheetName: 'Main',
  sheets: [
    {
      name: 'Main',
      usedRange: 'A1:A1',
      rowCount: 1,
      columnCount: 1,
      values: [['']],
      formulas: [['']],
      numberFormats: [['General']],
      structure: 'unknown',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

const createSheets: SubTask = {
  id: 's1',
  description: 'Create the month sheets',
  targetSheet: 'January',
  dependsOn: [],
  estimatedActions: 1,
};

const populate: SubTask = {
  id: 's2',
  description: 'Populate January headers',
  targetSheet: 'January',
  dependsOn: ['s1'],
  estimatedActions: 1,
};

describe('AgenticLoopService.runWave — stepwise execution', () => {
  let executor: jest.Mocked<Pick<ExecutorAgent, 'execute' | 'retryStep'>>;
  let verifier: jest.Mocked<Pick<VerifierAgent, 'verify'>>;
  let service: AgenticLoopService;
  let emitted: string[];

  const emitter = () => new SseEmitter((event: string) => { emitted.push(event); });

  beforeEach(() => {
    emitted = [];
    executor = { execute: jest.fn(), retryStep: jest.fn() };
    executor.retryStep.mockImplementation(async (retryContext, context, previousActions) =>
      executor.execute(retryContext.originalStep, context, previousActions),
    );
    verifier = {
      verify: jest.fn().mockResolvedValue({
        passed: true,
        feedback: 'ok',
        issues: [],
        subtaskResults: [],
      }),
    };

    const formulaAnalyzer = {
      analyzeSheet: jest.fn().mockReturnValue({ llmSummary: '' }),
    } as unknown as FormulaAnalyzer;
    const formulaValidator = {
      validatePreApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'pre_apply' }),
      checkPostApply: jest.fn().mockReturnValue({ passed: true, issues: [], phase: 'post_apply' }),
      formatFeedback: jest.fn().mockReturnValue(''),
      summarizeForVerifier: jest.fn().mockReturnValue('all checks passed'),
    } as unknown as FormulaValidatorService;
    const toolBridge = { waitForRangeData: jest.fn() } as unknown as ToolBridgeService;

    service = new AgenticLoopService(
      executor as unknown as ExecutorAgent,
      verifier as unknown as VerifierAgent,
      formulaAnalyzer,
      formulaValidator,
      toolBridge,
      new CompletenessChecker(),
      new FormattingChecker(),
      new OverwriteOccupancyChecker(),
    );
  });

  it('executes only the wave it is given, not the whole plan', async () => {
    const addSheet: Action = { type: 'ADD_SHEET', sheetName: 'January' } as Action;
    executor.execute.mockResolvedValue({ subtaskId: 's1', actions: [addSheet], isDone: true });

    const result = await service.runWave(
      'build a ledger',
      [createSheets],
      [],
      baseContext,
      emitter(),
    );

    // The second subtask exists in the plan but is NOT in this wave — no
    // Executor call may have been made for it (SD-3: no look-ahead).
    expect(executor.execute).toHaveBeenCalledTimes(1);
    expect(executor.execute.mock.calls[0][0].id).toBe('s1');
    expect(result.completedSubtasks.map((entry) => entry.subtaskId)).toEqual(['s1']);
  });

  it('reports only its own actions, never the prior waves it was seeded with', async () => {
    const priorAction: Action = { type: 'ADD_SHEET', sheetName: 'January' } as Action;
    const ownAction: Action = {
      type: 'SET_CELL',
      sheetName: 'January',
      row: 0,
      col: 0,
      value: 'Date',
    } as Action;

    executor.execute.mockResolvedValue({ subtaskId: 's2', actions: [ownAction], isDone: true });

    const result = await service.runWave(
      'build a ledger',
      [populate],
      [{ subtask: createSheets, actions: [priorAction] }],
      baseContext,
      emitter(),
    );

    // Double-emitting the prior wave's actions would re-apply work the user
    // already accepted — the whole hazard of carrying prior state.
    expect(result.actions).toEqual([ownAction]);
    expect(result.completedSubtasks.map((entry) => entry.subtaskId)).toEqual(['s2']);
  });

  it('makes prior waves visible to the Executor so it plans against created sheets', async () => {
    const priorAction: Action = { type: 'ADD_SHEET', sheetName: 'January' } as Action;
    executor.execute.mockResolvedValue({ subtaskId: 's2', actions: [], isDone: true });

    await service.runWave(
      'build a ledger',
      [populate],
      [{ subtask: createSheets, actions: [priorAction] }],
      baseContext,
      emitter(),
    );

    // The context handed to the Executor must contain the sheet the prior wave
    // created; without it, "populate January" plans against a workbook where
    // January does not exist and the Executor logs "Target sheet not found".
    const contextArg = executor.execute.mock.calls[0][1];
    expect(contextArg.sheets.map((sheet) => sheet.name)).toContain('January');
  });

  // Live incident (Sept 8, 2026): a stepwise wave that timed out with zero
  // actions emitted a hard `ERROR` SSE event — the same one the ONE-SHOT path
  // uses to mean "the whole request is over." The frontend treats `error` as
  // terminal (aborts the turn, stops the spinner) the instant it arrives, but
  // `executeStepwiseWave` treats an empty timed-out wave as RECOVERABLE — it
  // marks the wave skipped and tries the next one. The user saw "Agentic loop
  // timeout" and gave up on a turn the backend kept working on for several
  // more minutes underneath. Fixed by suppressing the ERROR emission
  // specifically for `runWave` (stepwise) callers.
  describe('timeout event suppression for stepwise waves', () => {
    it('does NOT emit a hard error event when a stepwise wave times out with nothing to show', async () => {
      // TIMEOUT_MS is a private readonly field — set it negative so the very
      // FIRST elapsed-time check (before any executor call) already trips,
      // giving a deterministic, instant timeout with zero actions produced.
      (service as unknown as { TIMEOUT_MS: number }).TIMEOUT_MS = -1;
      executor.execute.mockResolvedValue({ subtaskId: 's1', actions: [], isDone: false });

      const result = await service.runWave(
        'build a ledger',
        [createSheets],
        [],
        baseContext,
        emitter(),
      );

      expect(emitted).not.toContain('error');
      expect(result.actions).toEqual([]);
    });

    it('DOES emit a hard error event for the one-shot path under the identical timeout', async () => {
      (service as unknown as { TIMEOUT_MS: number }).TIMEOUT_MS = -1;
      executor.execute.mockResolvedValue({ subtaskId: 's1', actions: [], isDone: false });

      // `run()` is the pre-existing one-shot entry point — no `isStepwiseWave`
      // flag, so the ERROR this loop sends on a truly empty timeout must still
      // reach the client exactly as it did before this fix.
      await service.run('build a ledger', [createSheets], baseContext, emitter());

      expect(emitted).toContain('error');
    });
  });

  it('passes prior actions to the Executor as previousActions', async () => {
    const priorAction: Action = { type: 'ADD_SHEET', sheetName: 'January' } as Action;
    executor.execute.mockResolvedValue({ subtaskId: 's2', actions: [], isDone: true });

    await service.runWave(
      'build a ledger',
      [populate],
      [{ subtask: createSheets, actions: [priorAction] }],
      baseContext,
      emitter(),
    );

    expect(executor.execute.mock.calls[0][2]).toEqual([priorAction]);
  });
});
