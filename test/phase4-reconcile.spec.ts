import { reconcileRun, describeReconcileGaps } from '../src/agents/utils/reconcile.util';
import { Action, SubTask } from '../src/agents/types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 4 acceptance:
 *  - "Inject a failure (drop 3 month sheets from a wave): reconcile restores
 *     them without user action."
 *  - "Reconcile is a no-op (0 actions) on a fully correct build."
 *  - Dangling references on Main are reported.
 */

const COLUMNS = ['Unit No', 'Guest', 'Check In', 'Check Out', 'Total Amount'];
const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const monthSubtask = (month: string, i: number): SubTask => ({
  id: `p2_s${i + 1}`,
  description: `Create sheet '${month}' with headers in row 1 (A1:E1): ${COLUMNS.join(', ')}.`,
  targetSheet: month,
  dependsOn: [],
  estimatedActions: 3,
  expectedHeaders: COLUMNS,
});

/** What a correctly-built month sheet's actions look like. */
const builtMonth = (month: string): Action[] =>
  [
    { type: 'ADD_SHEET', name: month, sheetName: month },
    {
      type: 'BATCH_SET',
      sheetName: month,
      operations: COLUMNS.map((value, i) => ({
        address: `${String.fromCharCode(65 + i)}1`,
        value,
      })),
    },
    { type: 'CREATE_TABLE', sheetName: month, range: 'A1:E2', tableName: `tbl${month}`, hasHeaders: true },
    // The "Total Amount" column's formula. Without it this fixture describes
    // the TASKS.md #298 shape — headers right, column permanently blank — and
    // calling that "correctly built" is what let the live April sheet ship.
    { type: 'SET_FORMULA', sheetName: month, row: 1, col: 4, formula: '=D2-C2' },
  ] as unknown as Action[];

describe('reconcileRun (Phase 4, TASKS.md #287)', () => {
  const subtasks = MONTHS.map(monthSubtask);

  it('is a complete no-op on a fully correct build', () => {
    const applied = MONTHS.flatMap(builtMonth);
    const { gaps, repairActions } = reconcileRun({ subtasks, appliedActions: applied });

    expect(gaps).toEqual([]);
    expect(repairActions).toEqual([]);
  });

  it('restores three dropped month sheets WITHOUT user action (the acceptance case)', () => {
    const dropped = ['April', 'July', 'November'];
    const applied = MONTHS.filter((m) => !dropped.includes(m)).flatMap(builtMonth);

    const { gaps, repairActions } = reconcileRun({ subtasks, appliedActions: applied });

    expect(gaps.map((g) => g.sheet).sort()).toEqual([...dropped].sort());
    expect(gaps.every((g) => g.kind === 'missing-sheet' && g.repairable)).toBe(true);

    // Each restored month gets its sheet, its real headers and its table.
    for (const month of dropped) {
      const forMonth = repairActions.filter(
        (a) => ((a as { sheetName?: string }).sheetName ?? (a as { name?: string }).name) === month,
      );
      expect(forMonth.map((a) => a.type)).toEqual(['ADD_SHEET', 'BATCH_SET', 'CREATE_TABLE']);
      const headerOps = (
        forMonth[1] as unknown as { operations: Array<{ address: string; value: string }> }
      ).operations.filter((op) => /[A-Z]1$/.test(op.address));
      expect(headerOps.map((op) => op.value)).toEqual(COLUMNS);
    }
  });

  it('catches a sheet that exists but never got its header row (the Column1..N shape)', () => {
    const applied = [
      { type: 'ADD_SHEET', name: 'January', sheetName: 'January' },
      { type: 'CREATE_TABLE', sheetName: 'January', range: 'A1:E2', tableName: 'tblJanuary', hasHeaders: true },
    ] as unknown as Action[];

    const { gaps, repairActions } = reconcileRun({ subtasks: [monthSubtask('January', 0)], appliedActions: applied });

    expect(gaps).toHaveLength(1);
    expect(gaps[0].kind).toBe('missing-header-row');
    expect(gaps[0].repairable).toBe(true);
    const ops = (repairActions[0] as unknown as { operations: Array<{ value: string }> }).operations;
    expect(ops.map((op) => op.value)).toEqual(COLUMNS);
  });

  it('flags a header row that does not match the user’s columns, and does NOT guess a repair', () => {
    const applied = [
      { type: 'ADD_SHEET', name: 'January', sheetName: 'January' },
      {
        type: 'BATCH_SET',
        sheetName: 'January',
        operations: [
          { address: 'A1', value: 'Unit No' },
          { address: 'B1', value: 'Something Else' },
        ],
      },
    ] as unknown as Action[];

    const { gaps, repairActions } = reconcileRun({ subtasks: [monthSubtask('January', 0)], appliedActions: applied });

    const mismatch = gaps.find((g) => g.kind === 'header-mismatch');
    expect(mismatch).toBeDefined();
    expect(mismatch!.detail).toContain('Guest');
    expect(mismatch!.repairable).toBe(false);
    expect(repairActions).toEqual([]);
  });

  it('reports Main formulas pointing at sheets that were never created — and never invents a fix', () => {
    const main: SubTask = {
      id: 'p3_s1', description: 'Build Main', targetSheet: 'Main', dependsOn: [], estimatedActions: 5,
    };
    const applied = [
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      {
        type: 'BATCH_SET',
        sheetName: 'Main',
        operations: [{ address: 'B5', formula: "=SUM('February'!E:E)" }],
      },
    ] as unknown as Action[];

    const { gaps, repairActions } = reconcileRun({ subtasks: [main], appliedActions: applied });

    const dangling = gaps.find((g) => g.kind === 'dangling-reference');
    expect(dangling).toBeDefined();
    expect(dangling!.detail).toContain('February');
    expect(dangling!.repairable).toBe(false);
    expect(repairActions).toEqual([]);
  });

  it('counts a sheet that already existed before the run as present', () => {
    const { gaps } = reconcileRun({
      subtasks: [monthSubtask('January', 0)],
      appliedActions: [],
      preExistingSheets: ['January'],
    });
    expect(gaps.filter((g) => g.kind === 'missing-sheet')).toEqual([]);
  });

  it('cannot rebuild a missing sheet with no column list, and says so instead of pretending', () => {
    const plain: SubTask = {
      id: 's1', description: 'Create Summary', targetSheet: 'Summary', dependsOn: [], estimatedActions: 1,
    };
    const { gaps, repairActions } = reconcileRun({ subtasks: [plain], appliedActions: [] });

    expect(gaps[0].kind).toBe('missing-sheet');
    expect(gaps[0].repairable).toBe(false);
    expect(repairActions).toEqual([]);
  });
});

describe('describeReconcileGaps', () => {
  it('says nothing when there is nothing to say', () => {
    expect(describeReconcileGaps([])).toBe('');
  });

  it('separates what it fixed from what still needs attention', () => {
    const text = describeReconcileGaps([
      { kind: 'missing-sheet', sheet: 'April', detail: 'x', repairable: true },
      { kind: 'dangling-reference', sheet: 'Main', detail: 'Formulas on "Main" read from July.', repairable: false },
    ]);
    expect(text).toContain('April');
    expect(text).toContain('still need attention');
    expect(text).toContain('Main');
  });
});

/**
 * TASKS.md #298 — a subtask that reports success having skipped its own
 * formulas.
 *
 * Live shape: a smoke run built all twelve months with correct headers, and
 * April came back with 3 DATA_VALIDATION, 5 FORMAT_RANGE and 13
 * SET_COLUMN_WIDTH actions and ZERO formulas, `completed: true`, no
 * failedReason, while its eleven identical siblings all wrote theirs. Every
 * net passed it: the sheet exists, the header row is right, nothing dangles.
 * The user is promised a "total amount" column that can only ever be blank,
 * and the build says done — the exact silent failure §6 item 3 of
 * LONG_PROMPT_RELIABILITY_PLAN.md calls the one that actually matters.
 */
describe('Phase 4 reconcile — a derived column with no formula (TASKS.md #298)', () => {
  const HEADERS = ['Unit No', 'Guest', 'Check In', 'Check Out', 'Nights', 'Total Amount'];

  const builtSheet = (sheet: string): Action[] => [
    { type: 'ADD_SHEET', name: sheet, sheetName: sheet } as Action,
    {
      type: 'BATCH_SET',
      sheetName: sheet,
      operations: HEADERS.map((value, i) => ({
        address: `${String.fromCharCode(65 + i)}1`,
        value,
      })),
    } as Action,
  ];

  const restStep = (sheet: string): SubTask => ({
    id: `p2_${sheet}`,
    targetSheet: sheet,
    dependsOn: [],
    estimatedActions: 20,
    description: `Add formulas, validation and widths to '${sheet}'`,
    resolvedHeaderRow: HEADERS,
  });

  it('reports the April shape: formats and widths written, no formula anywhere', () => {
    const appliedActions: Action[] = [
      ...builtSheet('April'),
      { type: 'DATA_VALIDATION', sheetName: 'April', range: 'I2:I500', values: ['Paid'] } as Action,
      { type: 'FORMAT_RANGE', sheetName: 'April', range: 'A1:M1', bold: true } as Action,
      { type: 'SET_COLUMN_WIDTH', sheetName: 'April', col: 0, width: 80 } as Action,
    ];

    const { gaps } = reconcileRun({ subtasks: [restStep('April')], appliedActions });
    const derived = gaps.filter((g) => g.kind === 'derived-column-no-formula');

    expect(derived).toHaveLength(1);
    expect(derived[0].sheet).toBe('April');
    expect(derived[0].detail).toContain('Total Amount');
    expect(derived[0].repairable).toBe(false);
  });

  it('a sibling that DID write its formula is not reported', () => {
    const appliedActions: Action[] = [
      ...builtSheet('March'),
      { type: 'SET_FORMULA', sheetName: 'March', row: 1, col: 5, formula: '=E2*D2' } as Action,
    ];

    const { gaps } = reconcileRun({ subtasks: [restStep('March')], appliedActions });
    expect(gaps.filter((g) => g.kind === 'derived-column-no-formula')).toHaveLength(0);
  });

  it('a formula inside a BATCH_SET counts, and FILL_DOWN counts', () => {
    const viaBatch: Action[] = [
      ...builtSheet('May'),
      {
        type: 'BATCH_SET',
        sheetName: 'May',
        operations: [{ address: 'F2', formula: '=E2*D2' }],
      } as Action,
    ];
    const viaFill: Action[] = [
      ...builtSheet('June'),
      { type: 'FILL_DOWN', sheetName: 'June', range: 'F2:F100' } as Action,
    ];

    expect(
      reconcileRun({ subtasks: [restStep('May')], appliedActions: viaBatch }).gaps
        .filter((g) => g.kind === 'derived-column-no-formula'),
    ).toHaveLength(0);
    expect(
      reconcileRun({ subtasks: [restStep('June')], appliedActions: viaFill }).gaps
        .filter((g) => g.kind === 'derived-column-no-formula'),
    ).toHaveLength(0);
  });

  it('a sheet of pure input columns is never reported', () => {
    // The false positive that would matter: "Guest", "Bank Account" and
    // "Rate Per Night" are things the user types, not derivations.
    const inputsOnly: SubTask = {
      ...restStep('Lists'),
      resolvedHeaderRow: ['Guest', 'Bank Account', 'Rate Per Night', 'Source'],
    };
    const appliedActions: Action[] = [
      { type: 'ADD_SHEET', name: 'Lists', sheetName: 'Lists' } as Action,
    ];

    const { gaps } = reconcileRun({ subtasks: [inputsOnly], appliedActions });
    expect(gaps.filter((g) => g.kind === 'derived-column-no-formula')).toHaveLength(0);
  });

  it('the deterministic header step is exempt — formulas are not its job', () => {
    const headerStep: SubTask = {
      ...restStep('July'),
      id: 'hdr_July',
      isDeterministicHeaderStep: true,
    };

    const { gaps } = reconcileRun({
      subtasks: [headerStep],
      appliedActions: builtSheet('July'),
    });
    expect(gaps.filter((g) => g.kind === 'derived-column-no-formula')).toHaveLength(0);
  });

  it('the gap reaches the user in the completion message', () => {
    const { gaps } = reconcileRun({
      subtasks: [restStep('April')],
      appliedActions: builtSheet('April'),
    });
    expect(describeReconcileGaps(gaps)).toContain('April');
  });
});

/**
 * TASKS.md #310 — repair a missing derived formula by COPYING a sibling's.
 *
 * The April case again: eleven months carried `=E2*D2` in Total Amount and
 * April carried nothing. #300 made that visible; reporting it still leaves the
 * user with a permanently blank column and a manual fix. Eleven sheets having
 * built the identical column means the twelfth is not being invented — it is
 * being copied, which is the same reasoning that lets Phase 2 clone a
 * template's accepted actions.
 *
 * The line that must not be crossed: with NO sibling to copy from, nothing is
 * written. Inventing a formula there would turn a visible gap into a silently
 * wrong number, which is the one trade this codebase never makes.
 */
describe('Phase 4 reconcile — repairing a derived column from a sibling (TASKS.md #310)', () => {
  const HEADERS = ['Unit No', 'Guest', 'Check In', 'Check Out', 'Nights', 'Total Amount'];

  const builtSheet = (sheet: string): Action[] => [
    { type: 'ADD_SHEET', name: sheet, sheetName: sheet } as Action,
    {
      type: 'BATCH_SET',
      sheetName: sheet,
      operations: HEADERS.map((value, i) => ({
        address: `${String.fromCharCode(65 + i)}1`,
        value,
      })),
    } as Action,
  ];

  const withFormula = (sheet: string): Action[] => [
    ...builtSheet(sheet),
    { type: 'SET_FORMULA', sheetName: sheet, row: 1, col: 5, formula: '=E2*D2' } as Action,
  ];

  const restStep = (sheet: string): SubTask => ({
    id: `p2_${sheet}`,
    targetSheet: sheet,
    dependsOn: [],
    estimatedActions: 20,
    description: `Add formulas to '${sheet}'`,
    resolvedHeaderRow: HEADERS,
  });

  it('copies the formula its sibling sheets already use', () => {
    const appliedActions = [
      ...withFormula('March'),
      ...withFormula('May'),
      ...builtSheet('April'), // the one that skipped its formulas
    ];

    const { gaps, repairActions } = reconcileRun({
      subtasks: [restStep('March'), restStep('April'), restStep('May')],
      appliedActions,
    });

    const repair = repairActions.find(
      (a) => (a as unknown as Record<string, unknown>).sheetName === 'April',
    ) as unknown as Record<string, unknown>;
    expect(repair).toBeDefined();
    expect(repair.type).toBe('SET_FORMULA');
    expect(repair.formula).toBe('=E2*D2');
    // Same column and row as the sibling, or the relative refs would be wrong.
    expect(repair.col).toBe(5);
    expect(repair.row).toBe(1);

    const gap = gaps.find((g) => g.kind === 'derived-column-no-formula');
    expect(gap?.repairable).toBe(true);
    expect(gap?.sheet).toBe('April');
  });

  it('writes NOTHING when there is no sibling to copy from', () => {
    // The line. A lone sheet missing its formula is reported, never invented.
    const { gaps, repairActions } = reconcileRun({
      subtasks: [restStep('April')],
      appliedActions: builtSheet('April'),
    });

    expect(repairActions).toEqual([]);
    const gap = gaps.find((g) => g.kind === 'derived-column-no-formula');
    expect(gap?.repairable).toBe(false);
    expect(gap?.detail).toContain('no other sheet built that column');
  });

  it('refuses a sibling whose column sits in a DIFFERENT position', () => {
    // A formula written for column F cannot be dropped into column D: its own
    // relative references would land on the wrong data.
    const shifted = ['Guest', 'Total Amount', 'Check In'];
    const sibling: Action[] = [
      { type: 'ADD_SHEET', name: 'March', sheetName: 'March' } as Action,
      {
        type: 'BATCH_SET',
        sheetName: 'March',
        operations: shifted.map((value, i) => ({
          address: `${String.fromCharCode(65 + i)}1`,
          value,
        })),
      } as Action,
      { type: 'SET_FORMULA', sheetName: 'March', row: 1, col: 1, formula: '=A2*1' } as Action,
    ];

    const { repairActions } = reconcileRun({
      subtasks: [
        { ...restStep('March'), resolvedHeaderRow: shifted },
        restStep('April'),
      ],
      appliedActions: [...sibling, ...builtSheet('April')],
    });

    expect(repairActions).toEqual([]);
  });

  it('a sheet that already has its formula is left untouched', () => {
    const { gaps, repairActions } = reconcileRun({
      subtasks: [restStep('March'), restStep('April')],
      appliedActions: [...withFormula('March'), ...withFormula('April')],
    });

    expect(repairActions).toEqual([]);
    expect(gaps.filter((g) => g.kind === 'derived-column-no-formula')).toHaveLength(0);
  });
});
