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

  /**
   * Live incident (TASKS.md #274): after #271 folded a whole month's build
   * (create + headers + table + row-2 formulas + dropdowns + widths + font)
   * into ONE subtask, that subtask needs enough Executor iterations to do all
   * of it — and one that dies partway (hit the 10-iteration cap after
   * ADD_SHEET + a headerless CREATE_TABLE, but before ever writing the header
   * row) still shipped those two actions as if the step had succeeded. The
   * user saw "✓ Applied" and a table with Excel's own placeholder headers
   * ("Column1"..."Column13") — a workbook that LOOKED built but was garbage.
   *
   * Root cause: `buildLoopResult`'s legacy fallback branch (only reached when
   * NOTHING in the wave safely completed) used to flatten every subtask's
   * accumulated actions with no check that the subtask itself ever finished.
   */
  it('ships NOTHING for a subtask that dies mid-way, rather than its partial fragments (TASKS.md #274)', async () => {
    const subtasks: SubTask[] = [
      {
        id: 's1',
        description: "Create sheet 'December' and build the booking/payment log",
        targetSheet: 'December',
        dependsOn: [],
        estimatedActions: 8,
      },
    ];

    // ADD_SHEET lands, then CREATE_TABLE lands (with no header row ever
    // written first — exactly the live shape), then the subtask never
    // reaches isDone and exhausts the iteration cap.
    let call = 0;
    executor.execute.mockImplementation(async (subtask: SubTask) => {
      call += 1;
      if (call === 1) {
        return {
          subtaskId: subtask.id,
          actions: [{ type: 'ADD_SHEET', sheetName: 'December' } as Action],
          isDone: false,
        };
      }
      if (call === 2) {
        return {
          subtaskId: subtask.id,
          actions: [
            {
              type: 'CREATE_TABLE',
              sheetName: 'December',
              range: 'A1:M1',
              tableName: 'tblDecember',
              hasHeaders: true,
            } as Action,
          ],
          isDone: false,
        };
      }
      return { subtaskId: subtask.id, actions: [], isDone: false, nextStep: 'still working' };
    });

    const result = await service.run(
      'Build the December sheet',
      subtasks,
      baseContext,
      new SseEmitter(emit),
    );

    // The wave itself is judged "low-risk, skip verification" (2 harmless,
    // non-destructive, non-formula actions) — that decision is correct in
    // isolation and is not what this test is about; #274 is that it must
    // not ALSO resurrect the subtask that produced them as "completed".
    expect(result.completedSubtasks).toEqual([]);
    // The actual regression: this used to be [ADD_SHEET, CREATE_TABLE] —
    // a headerless table shipped and applied as if the step had succeeded.
    expect(result.actions).toEqual([]);
    // And the failure must be REPORTED, not just silently excluded — a
    // dropped step with no reason anywhere is its own false-completeness
    // shape (the user sees "Applied" with nothing to explain what changed).
    expect(result.failedSubtasks.some((f) => f.subtaskId === 's1')).toBe(true);
    expect(result.failedSubtask?.reason).toMatch(/max iterations/i);
  });

  /**
   * Live incident (TASKS.md #275): two consecutive real runs of the same
   * 12-month-sheet prompt each had EVERY ONE of the 12 independent
   * month-create subtasks — grouped into a single parallel wave since they
   * have no dependsOn between them — fail to reach `completed: true`, with a
   * different subset each run naming `OpenRouter could not verify available
   * credits for this request in time` and the rest exhausting
   * MAX_ITERATIONS_PER_SUBTASK. The signature (different random subset
   * failing each run, not the same ones) pointed at 12 simultaneous Executor
   * LLM calls overwhelming the provider's own concurrency ceiling, not 12
   * independent bugs. `executeSubtask` was fired via one unthrottled
   * `Promise.all(wave.map(...))` with no cap on how many ran at once.
   */
  /**
   * Phase 2 of LONG_PROMPT_RELIABILITY_PLAN.md: 12 month subtasks that are the
   * same work on different sheets build ONE sheet with the Executor; the other
   * 11 are stamped from its accepted actions.
   */
  describe('template + replicate (Phase 2)', () => {
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December',
    ];
    const monthSubtasks = (): SubTask[] =>
      months.map((month, i) => ({
        id: `p2_s${i + 1}`,
        description: `Create sheet '${month}' and build the log. Table 'tbl${month.slice(0, 3)}'.`,
        targetSheet: month,
        dependsOn: [],
        estimatedActions: 2,
      }));
    const passAll = () =>
      verifier.verify.mockImplementation(async (_prompt, verifySubtasks) => ({
        passed: true,
        feedback: 'ok',
        issues: [],
        subtaskResults: (verifySubtasks as SubTask[]).map((s) => ({
          subtaskId: s.id, passed: true, feedback: 'OK', issues: [],
        })),
      }));
    const sheetNamesOf = (actions: Action[]) =>
      actions.map((a) => (a as { sheetName?: string }).sheetName);

    it('makes ONE Executor generation for 12 identical-shape month sheets and delivers all 12', async () => {
      executor.execute.mockImplementation(async (subtask: SubTask) => ({
        subtaskId: subtask.id,
        actions: [
          { type: 'ADD_SHEET', sheetName: subtask.targetSheet } as Action,
          { type: 'CREATE_TABLE', sheetName: subtask.targetSheet, range: 'A1:M2', tableName: `tbl${subtask.targetSheet.slice(0, 3)}`, hasHeaders: true } as Action,
        ],
        isDone: true,
      }));
      passAll();

      const result = await service.run('Create 12 month sheets', monthSubtasks(), baseContext, new SseEmitter(emit));

      expect(executor.execute).toHaveBeenCalledTimes(1);
      expect(result.failedSubtasks).toEqual([]);
      const delivered = result.actions;
      expect(new Set(sheetNamesOf(delivered.filter((a) => a.type === 'ADD_SHEET')))).toEqual(new Set(months));
      const tables = delivered.filter((a) => a.type === 'CREATE_TABLE') as Array<{ sheetName: string; tableName: string }>;
      expect(tables.find((t) => t.sheetName === 'February')?.tableName).toBe('tblFeb');
      expect(tables.find((t) => t.sheetName === 'December')?.tableName).toBe('tblDec');
    });

    it('falls back to building each sibling itself when the template never completes', async () => {
      executor.execute.mockImplementation(async (subtask: SubTask) =>
        subtask.id === 'p2_s1'
          ? { subtaskId: subtask.id, actions: [], isDone: false, nextStep: 'still working' }
          : {
              subtaskId: subtask.id,
              actions: [{ type: 'ADD_SHEET', sheetName: subtask.targetSheet } as Action],
              isDone: true,
            },
      );
      passAll();

      const result = await service.run('Create 12 month sheets', monthSubtasks(), baseContext, new SseEmitter(emit));

      expect(result.failedSubtasks.map((f) => f.subtaskId)).toEqual(['p2_s1']);
      expect(new Set(sheetNamesOf(result.actions))).toEqual(new Set(months.slice(1)));
      const siblingCalls = executor.execute.mock.calls.filter((c) => (c[0] as SubTask).id !== 'p2_s1');
      expect(siblingCalls).toHaveLength(11); // each built individually, none cloned from a failed template
    });
  });

  it('never runs more than MAX_WAVE_CONCURRENCY subtasks at once within a wave (TASKS.md #275)', async () => {
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December',
    ];
    const subtasks: SubTask[] = months.map((month, i) => ({
      id: `s${i + 1}`,
      // Distinct work per sheet, so Phase 2 cloning does not apply and the cap
      // itself is what is under test.
      description: `Create ${month} sheet (variant ${i})`,
      targetSheet: month,
      dependsOn: [],
      estimatedActions: 1,
    }));

    let inFlight = 0;
    let maxInFlight = 0;
    executor.execute.mockImplementation(async (subtask: SubTask) => {
      inFlight += 1;
      maxInFlight = Math.max(maxInFlight, inFlight);
      await new Promise((resolve) => setTimeout(resolve, 5));
      inFlight -= 1;
      return {
        subtaskId: subtask.id,
        actions: [{ type: 'ADD_SHEET', sheetName: subtask.targetSheet } as Action],
        isDone: true,
      };
    });
    verifier.verify.mockResolvedValue({
      passed: true,
      feedback: 'ok',
      issues: [],
      subtaskResults: subtasks.map((s) => ({
        subtaskId: s.id,
        passed: true,
        feedback: 'OK',
        issues: [],
      })),
    });

    await service.run('Create 12 month sheets', subtasks, baseContext, new SseEmitter(emit));

    // Phase 3 (TASKS.md #286) made this width adaptive: it STARTS at 3 and
    // earns its way up to the ceiling after runs of clean successes, which is
    // exactly what happens here since nothing fails. The invariant that
    // matters — and the one the live incident was about — is that a 12-subtask
    // wave never fires all 12 at once.
    expect(maxInFlight).toBeLessThanOrEqual(6); // the controller's ceiling
    expect(maxInFlight).toBeLessThan(months.length); // never the whole wave
    expect(maxInFlight).toBeGreaterThan(1); // still parallel, just throttled
  });

  /**
   * Live incident (TASKS.md #278): a real run had January's template time out.
   * `buildLoopResult`'s FINAL result correctly excluded its actions (#274),
   * but the PROGRESSIVE per-wave `onWaveComplete` callback — used to render an
   * "Applied" Accept card mid-build, before the run finishes — flattened every
   * subtask's actions in the wave with no completion check at all. May and
   * July's partial in-progress header cells (3 and 4 columns, from a subtask
   * that never reached `completed: true`) were shown as "Applied" through
   * this path even though the final result never contained them.
   */
  it('never hands onWaveComplete a subtask that did not reach completed:true (TASKS.md #278)', async () => {
    const subtasks: SubTask[] = [
      { id: 'may', description: 'Create May sheet', targetSheet: 'May', dependsOn: [], estimatedActions: 8 },
      { id: 'august', description: 'Create August sheet', targetSheet: 'August', dependsOn: [], estimatedActions: 8 },
    ];

    // May: partial header cells land, then it never reaches isDone (dies mid-way).
    // August: finishes cleanly.
    executor.execute.mockImplementation(async (subtask: SubTask) => {
      if (subtask.id === 'may') {
        return {
          subtaskId: subtask.id,
          actions: [
            { type: 'ADD_SHEET', sheetName: 'May' } as Action,
            {
              type: 'BATCH_SET', sheetName: 'May',
              operations: [{ address: 'A1', value: 'Unit No' }, { address: 'B1', value: 'Guest' }, { address: 'C1', value: 'Check In' }],
            } as Action,
          ],
          isDone: false,
          nextStep: 'still working',
        };
      }
      return {
        subtaskId: subtask.id,
        actions: [{ type: 'ADD_SHEET', sheetName: 'August' } as Action],
        isDone: true,
      };
    });
    verifier.verify.mockResolvedValue({
      passed: true, feedback: 'ok', issues: [],
      subtaskResults: [{ subtaskId: 'august', passed: true, feedback: 'OK', issues: [] }],
    });

    const waveCards: Action[][] = [];
    await service.run('Build May and August', subtasks, baseContext, new SseEmitter(emit), {
      onWaveComplete: async (actions) => {
        waveCards.push(actions);
      },
    });

    expect(waveCards).toHaveLength(1);
    const cardSheets = new Set(waveCards[0].map((a) => (a as { sheetName?: string }).sheetName));
    expect(cardSheets).toEqual(new Set(['August']));
    expect(cardSheets.has('May')).toBe(false);
  });
});
