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
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Live incident (TASKS.md #281): "stop button is not active there, but it's
 * still thinking" — the client's Stop click aborts its own fetch, but a
 * single heavy subtask (still mid-iteration inside `executeSubtask`) kept
 * calling the Executor for every remaining iteration regardless. `abortSignal`
 * was only ever checked BETWEEN waves (line ~303) and in the verify/retry
 * cycle (line ~596) — never inside the per-subtask iteration loop itself,
 * which is exactly where a long build actually spends most of its time.
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

describe('AgenticLoopService — abort mid-iteration (TASKS.md #281)', () => {
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

  it('stops calling the Executor once aborted, even mid-way through one subtask’s own iterations', async () => {
    const subtask: SubTask = {
      id: 's1',
      description: 'Build a heavy sheet',
      targetSheet: 'January',
      dependsOn: [],
      estimatedActions: 20,
    };

    const controller = new AbortController();
    let calls = 0;
    executor.execute.mockImplementation(async () => {
      calls += 1;
      // Abort partway through, exactly like a real user clicking Stop while
      // the subtask is still iterating (nothing ever hit isDone/timeout).
      if (calls === 3) controller.abort();
      return { subtaskId: subtask.id, actions: [], isDone: false, nextStep: 'still working' };
    });

    const result = await service.run(
      'Build January',
      [subtask],
      baseContext,
      new SseEmitter(emit),
      { abortSignal: controller.signal },
    );

    // The 4th call would have happened without the fix — iteration 3 sets
    // isDone:false and returns, the NEXT loop pass is where the abort check
    // must catch it before calling the Executor a 4th time.
    expect(calls).toBe(3);
    expect(result.failedSubtasks[0]?.reason).toMatch(/cancelled/i);
  });

  it('does not cancel a subtask that finishes before the signal ever fires', async () => {
    const subtask: SubTask = {
      id: 's1', description: 'Quick step', targetSheet: 'Lists', dependsOn: [], estimatedActions: 1,
    };
    const controller = new AbortController();
    executor.execute.mockResolvedValue({
      subtaskId: subtask.id,
      actions: [{ type: 'ADD_SHEET', name: 'Lists' } as never],
      isDone: true,
    });

    const result = await service.run(
      'Build Lists',
      [subtask],
      baseContext,
      new SseEmitter(emit),
      { abortSignal: controller.signal },
    );

    expect(result.failedSubtasks).toEqual([]);
    expect(result.actions).toHaveLength(1);
  });
});
