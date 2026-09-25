import { StructuralIntentChecker } from '../src/agents/checkers/structural-intent.checker';
import { OverwriteOccupancyChecker } from '../src/agents/checkers/overwrite-occupancy.checker';
import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { Action, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Checker context discipline — the single root cause behind THREE separate
 * live failures in one session (TASKS.md #294, #296, #302).
 *
 * Every one of them was a checker handed a WorkbookContext that could not
 * distinguish the run's own writes from the workbook the user actually has:
 *
 *   #294  a sheet merely WRITTEN to was read as created, because
 *         `virtualApply`'s `ensureSheet` conjures a shadow sheet on any write
 *         — so the check meant to catch "you never created this" was skipped.
 *   #296  a sheet genuinely CREATED in the batch was read as absent, because
 *         `shadowAsContext` only walked the pre-batch sheet list — so a
 *         subtask could not reference the sheet it was itself creating.
 *   #302  a cell occupied by the subtask's OWN un-accepted attempt was read as
 *         pre-existing data — so every retry was refused for having already
 *         done the work, and twelve month steps died in one run.
 *
 * Each was fixed in its own place. This file states the RULE the three share,
 * so a fourth instance fails here rather than in a workbook. The rule:
 *
 *   A checker asking "what was here BEFORE this work?" must be given a
 *   context without that work in it. A checker asking "is the RESULT right?"
 *   must be given a context with it. Neither may be handed the other's.
 *
 * These assertions are deliberately at the checker boundary rather than
 * through the loop: they describe the contract each checker's context must
 * satisfy, which is the thing that kept being got wrong.
 */

const baseSheet = (name: string, values: unknown[][]) => ({
  name,
  usedRange: `A1:${String.fromCharCode(64 + Math.max(1, values[0]?.length ?? 1))}${values.length}`,
  rowCount: values.length,
  columnCount: values[0]?.length ?? 0,
  values,
  formulas: values.map((row) => row.map(() => '')),
  numberFormats: values.map((row) => row.map(() => 'General')),
  structure: 'data_table' as const,
  headerRowIndex: 0,
});

const contextOf = (sheets: ReturnType<typeof baseSheet>[]): WorkbookContext => ({
  activeSheetName: sheets[0]?.name ?? 'Sheet1',
  sheets,
  namedRanges: [],
  tables: [],
});

const subtask = (over: Partial<SubTask> & { id: string }): SubTask => ({
  targetSheet: 'Main',
  dependsOn: [],
  estimatedActions: 5,
  description: 'Do the work',
  ...over,
});

describe('Checker context discipline — the #294 / #296 / #302 rule', () => {
  const emptyWorkbook = contextOf([baseSheet('Sheet1', [['']])]);

  describe('a "what was here before?" checker must NOT see the run\'s own writes', () => {
    it('#294: a sheet the subtask only WROTE to is not credited as created', () => {
      // The pre-run context is the one that can tell the difference. Handing
      // this checker a shadow-derived context (where the write conjured the
      // sheet) is what made the live run report success with Main missing.
      const checker = new StructuralIntentChecker();
      const result = checker.check(
        [
          {
            subtask: subtask({
              id: 'p3_s1',
              description: "Create sheet 'Main' positioned first and write the dashboard title.",
            }),
            actions: [
              { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'Dashboard' },
            ] as unknown as Action[],
          },
        ],
        emptyWorkbook,
        new Set<string>(),
      );

      expect(result.passed).toBe(false);
    });

    it('#302: a cell written by this very subtask is not "already occupied"', () => {
      // The context must exclude the subtask's own writes. If the formula it
      // just wrote is present, the checker refuses the subtask for its own
      // work and no retry can ever succeed.
      const checker = new OverwriteOccupancyChecker();
      const beforeOwnWrites = contextOf([
        baseSheet('January', [
          ['Unit No', 'Guest', 'Rate', 'Total Amount'],
          ['', '', '', ''],
        ]),
      ]);

      const result = checker.check(
        [
          {
            subtask: subtask({
              id: 'p2_s1',
              targetSheet: 'January',
              description:
                'Add only what is still needed: row-2 formulas, and set column widths A=80.',
            }),
            actions: [
              { type: 'SET_FORMULA', sheetName: 'January', row: 1, col: 3, formula: '=B2*C2' },
            ] as unknown as Action[],
          },
        ],
        beforeOwnWrites,
      );

      expect(result.passed).toBe(true);
    });

    it('#302 guard: a cell the USER already filled is still refused', () => {
      // The discipline must not become "never refuse anything".
      const checker = new OverwriteOccupancyChecker();
      const occupied = contextOf([
        baseSheet('January', [
          ['Unit No', 'Guest', 'Rate', 'Total Amount'],
          ['', '', '', 1200],
        ]),
      ]);

      const result = checker.check(
        [
          {
            subtask: subtask({
              id: 'p2_s1',
              targetSheet: 'January',
              description:
                'Add only what is still needed: row-2 formulas, and set column widths A=80.',
            }),
            actions: [
              { type: 'SET_FORMULA', sheetName: 'January', row: 1, col: 3, formula: '=B2*C2' },
            ] as unknown as Action[],
          },
        ],
        occupied,
      );

      expect(result.passed).toBe(false);
    });
  });

  describe('a "is the result right?" checker MUST see the run\'s own writes', () => {
    it('#296: a formula may reference the sheet its own batch creates', () => {
      const validator = new FormulaValidatorService();
      const result = validator.validatePreApply(
        [
          { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
          { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(B5:B16)' },
        ] as unknown as Action[],
        emptyWorkbook,
        'Main',
        buildShadowWorkbook(emptyWorkbook),
      );

      expect(result.issues.filter((i) => i.severity === 'error')).toEqual([]);
    });

    it('#296 guard: a sheet nothing in the batch touches is still unknown', () => {
      const validator = new FormulaValidatorService();
      const result = validator.validatePreApply(
        [
          { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
          { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(Nowhere!B5:B16)' },
        ] as unknown as Action[],
        emptyWorkbook,
        'Main',
        buildShadowWorkbook(emptyWorkbook),
      );

      expect(result.passed).toBe(false);
      expect(result.issues.some((i) => i.message.includes('unknown sheet "Nowhere"'))).toBe(true);
    });
  });

  /**
   * The rule stated as one assertion, so the NEXT checker added to
   * `runDeterministicChecks` has something to be measured against: the two
   * kinds of context must actually differ for a run that writes anything.
   * If they ever collapse into one, all three bugs above become possible
   * again simultaneously.
   */
  it('the two contexts are genuinely different objects for a run that writes', () => {
    const before = emptyWorkbook;
    const shadow = buildShadowWorkbook(before);
    const validator = new FormulaValidatorService();

    // A batch that both creates a sheet and writes to it.
    const actions = [
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'Dashboard' },
    ] as unknown as Action[];

    // The "result" view resolves Main; the "before" view does not know it.
    const afterView = validator.validatePreApply(
      [...actions, { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 0, formula: '=A1' } as unknown as Action],
      before,
      'Main',
      shadow,
    );
    expect(afterView.issues.some((i) => i.message.includes('unknown sheet'))).toBe(false);
    expect(before.sheets.some((s) => s.name === 'Main')).toBe(false);
  });
});
