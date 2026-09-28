import { StructuralIntentChecker } from '../src/agents/checkers/structural-intent.checker';
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';

describe('StructuralIntentChecker', () => {
  const checker = new StructuralIntentChecker();

  const emptyWorkbook: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      {
        name: 'Sheet1',
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

  const createMainSubtask: SubTask = {
    id: 's1',
    description: "Create sheet 'Main' if it doesn't exist",
    targetSheet: 'Main',
    dependsOn: [],
    estimatedActions: 1,
  };

  it('fails when CREATE_SHEET lands under a different name than the subtask targets (COMPETITIVE_STUDY_SHORTCUT.md:71 — Main -> Main 2)', () => {
    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          actions: [{ type: 'CREATE_SHEET', sheetName: 'Main 2' }],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(false);
    expect(result.subtaskResults[0].passed).toBe(false);
    expect(result.subtaskResults[0].feedback).toContain('Main');
    expect(result.subtaskResults[0].feedback).toContain('Main 2');
  });

  it('passes when CREATE_SHEET uses the exact requested name', () => {
    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          actions: [{ type: 'CREATE_SHEET', sheetName: 'Main' }],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(true);
  });

  it('passes when ADD_SHEET (the normalized type) uses the "name" field matching the target', () => {
    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          actions: [{ type: 'ADD_SHEET', name: 'Main' }],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(true);
  });

  it('skips the check entirely when the target sheet already exists (idempotent reuse is correct behavior)', () => {
    const contextWithMain: WorkbookContext = {
      ...emptyWorkbook,
      sheets: [...emptyWorkbook.sheets, { ...emptyWorkbook.sheets[0], name: 'Main' }],
    };

    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          // Executor correctly emitted nothing further, or referenced the existing sheet.
          actions: [],
        },
      ],
      contextWithMain,
    );

    expect(result.passed).toBe(true);
  });

  it('skips subtasks that are not about creating a sheet', () => {
    const result = checker.check(
      [
        {
          subtask: {
            id: 's2',
            description: 'Set Remarks to Cleared where Payment Status = Paid',
            targetSheet: 'Purchase Register',
            dependsOn: [],
            estimatedActions: 1,
          },
          actions: [{ type: 'SET_MATCHING_ROWS', targetColumn: 'Remarks' } as never],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(true);
    expect(result.subtaskResults[0].feedback).toContain('skipped');
  });

  it('ignores CREATE_SHEET actions belonging to other subtasks in the same batch', () => {
    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          actions: [{ type: 'CREATE_SHEET', sheetName: 'Main' }],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(true);
  });

  /**
   * TASKS.md #266 — live repro: a 12-subtask batched Executor call emitted
   * every OTHER month's ADD_SHEET but dropped December's outright, while
   * December's own writes (BATCH_SET, CREATE_TABLE, DATA_VALIDATION,
   * SET_COLUMN_WIDTH) all went through untouched. The subtask's own
   * description said "Create sheet 'December'. Write headers..." — this is
   * exactly the shape the checker existed to police, but the old loop only
   * ever compared an EXISTING create's name; zero matching creates fell
   * through with no issue raised at all.
   */
  describe('missing create entirely (TASKS.md #266)', () => {
    const createDecemberSubtask: SubTask = {
      id: 's12',
      description:
        "Create sheet 'December'. Write headers in A1:J1: Unit No, Guest, Guest Name, Check In, " +
        'Check Out, Rate Per Night, Total Amount, Source, Payment Status, Bank Account. ' +
        'Create Excel Table tblDec over A1:J2.',
      targetSheet: 'December',
      dependsOn: [],
      estimatedActions: 8,
    };

    it('fails when every action is a write and none of them is the create the description promised', () => {
      const result = checker.check(
        [
          {
            subtask: createDecemberSubtask,
            actions: [
              {
                type: 'BATCH_SET',
                sheetName: 'December',
                operations: [{ address: 'A1', value: 'Unit No' }],
              } as never,
              { type: 'CREATE_TABLE', sheetName: 'December', range: 'A1:J2', tableName: 'tblDec' } as never,
            ],
          },
        ],
        emptyWorkbook,
      );

      expect(result.passed).toBe(false);
      expect(result.subtaskResults[0].feedback).toContain('December');
      expect(result.subtaskResults[0].feedback).toMatch(/no ADD_SHEET\/CREATE_SHEET/);
    });

    it('fails on a completely empty action list too, not just a writes-only one', () => {
      const result = checker.check(
        [{ subtask: createDecemberSubtask, actions: [] }],
        emptyWorkbook,
      );
      expect(result.passed).toBe(false);
    });

    it('passes once the matching create is present alongside the writes', () => {
      const result = checker.check(
        [
          {
            subtask: createDecemberSubtask,
            actions: [
              { type: 'ADD_SHEET', name: 'December' } as never,
              {
                type: 'BATCH_SET',
                sheetName: 'December',
                operations: [{ address: 'A1', value: 'Unit No' }],
              } as never,
            ],
          },
        ],
        emptyWorkbook,
      );
      expect(result.passed).toBe(true);
    });

    it('does not double-report when a wrong-named create already failed the subtask', () => {
      const result = checker.check(
        [
          {
            subtask: createDecemberSubtask,
            actions: [{ type: 'ADD_SHEET', name: 'December 2' } as never],
          },
        ],
        emptyWorkbook,
      );
      // Exactly the pre-existing wrong-name issue — not ALSO a "missing create" one.
      expect(result.subtaskResults[0].issues).toHaveLength(1);
      expect(result.subtaskResults[0].feedback).toContain('December 2');
    });
  });
});

/**
 * Live incident (TASKS.md #294): a smoke run completed with Main written to
 * but never created. `p3_s1`'s description said "Create sheet 'Main'…", it
 * emitted the title write and no ADD_SHEET, and this checker — which exists
 * precisely to catch that — passed it.
 *
 * Cause: `virtualApply`'s `ensureSheet` conjures a sheet the moment anything
 * writes to it, so Main appeared in the shadow-enriched context, and the
 * "reusing an existing sheet is correct" guard skipped the check. The guard is
 * right; the context it was given was not.
 */
describe('StructuralIntentChecker — pre-run vs shadow context (TASKS.md #294)', () => {
  const checker = new StructuralIntentChecker();

  const subtask = {
    id: 'p3_s1',
    description: "Create sheet 'Main' positioned first with the title in A1 'Payments Dashboard'",
    targetSheet: 'Main',
    dependsOn: [],
    estimatedActions: 2,
  };

  /** The live shape: writes to Main, no create. */
  const actions = [
    { type: 'BATCH_SET', sheetName: 'Main', operations: [{ address: 'A1', value: 'Payments Dashboard' }] },
  ] as never[];

  const contextWith = (sheetNames: string[]): WorkbookContext => ({
    activeSheetName: 'Sheet1',
    sheets: sheetNames.map((name) => ({
      name, usedRange: 'A1:A1', rowCount: 1, columnCount: 1,
      values: [['']], formulas: [['']], numberFormats: [['General']],
      structure: 'data_table' as const, headerRowIndex: 0,
    })),
    namedRanges: [],
    tables: [],
  });

  it('FAILS the subtask when Main did not exist before the run (the pre-run context)', () => {
    const result = checker.check([{ subtask, actions }], contextWith(['Sheet1']));
    expect(result.passed).toBe(false);
    expect(result.subtaskResults[0].feedback).toMatch(/no ADD_SHEET\/CREATE_SHEET/i);
  });

  it('would have PASSED against the shadow context — the bug this documents', () => {
    // Main present only because the run's own write conjured it.
    const result = checker.check([{ subtask, actions }], contextWith(['Sheet1', 'Main']));
    expect(result.passed).toBe(true);
  });

  it('still skips a sheet that genuinely existed before the run', () => {
    const reuse = { ...subtask, description: "Create sheet 'Main' if it does not already exist" };
    const result = checker.check([{ subtask: reuse, actions }], contextWith(['Sheet1', 'Main']));
    expect(result.passed).toBe(true);
  });
});
