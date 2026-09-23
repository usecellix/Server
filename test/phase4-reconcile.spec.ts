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
