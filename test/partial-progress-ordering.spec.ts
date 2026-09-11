import { splitIntoActionWaves } from '../src/excel-ai/utils/action-wave.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #172 — a partially-completed build must still apply in a safe order.
 *
 * From a live run: Main!B11..B16 came back `#REF!` and C11..D16 `#VALUE!` —
 * rows 11-16 being exactly July..December. The formulas were correct
 * (`=SUM(July!H:H)`), the dependency graph was correct (`sM3` dependsOn
 * `sJul..sDec`), and the plan was correct. What was wrong was the order the
 * actions reached EXCEL.
 *
 * `planner.prompt.ts` rule 0 deliberately emits the Main-sheet subtasks FIRST
 * and the twelve month subtasks LAST, so a truncated plan loses boilerplate
 * rather than the dashboard. Right for planning, actively dangerous for
 * applying: it puts Main's cross-sheet formulas ahead of the ADD_SHEETs that
 * create the sheets they reference. The success path is saved by
 * `splitIntoActionWaves`' phase buckets; the partial-progress path built its
 * preview straight from `finalizeActions`, which does not reorder.
 *
 * This locks the invariant that fixes it, expressed the way the failure
 * presented: every sheet a formula names must be created earlier in the array.
 */
function planOrderedLikeRule0(): SheetActionPayload[] {
  const months = ['July', 'August', 'September', 'October', 'November', 'December'];

  // Main first — formulas referencing sheets that do not exist yet.
  const mainFirst: SheetActionPayload[] = [
    { type: 'ADD_SHEET', sheetName: 'Main' },
    ...months.map((m, i) => ({
      type: 'SET_FORMULA',
      sheetName: 'Main',
      row: 10 + i,
      col: 1,
      formula: `=SUM(${m}!H:H)`,
    })),
  ] as SheetActionPayload[];

  // Month sheets last, exactly as rule 0 requires.
  const monthsLast: SheetActionPayload[] = months.flatMap((m) => [
    { type: 'ADD_SHEET', sheetName: m },
    { type: 'BATCH_SET', sheetName: m, operations: [{ address: 'A1', value: 'Unit No' }] },
  ]) as SheetActionPayload[];

  // Padding so the batch clears MIN_ACTIONS_TO_STAGE and spans >1 phase.
  const padding: SheetActionPayload[] = Array.from({ length: 20 }, (_, i) => ({
    type: 'FORMAT_RANGE',
    sheetName: 'Main',
    row: i,
    col: 0,
    rowCount: 1,
    colCount: 1,
    format: { bold: true },
  })) as SheetActionPayload[];

  return [...mainFirst, ...monthsLast, ...padding];
}

/** Index of the first action that creates `sheet`, or -1. */
function createIndex(actions: SheetActionPayload[], sheet: string): number {
  return actions.findIndex(
    (a) =>
      (a.type === 'ADD_SHEET' || a.type === 'CREATE_SHEET') &&
      String(a.sheetName ?? a.name ?? '') === sheet,
  );
}

describe('#172 — phase ordering fixes the live #REF! failure', () => {
  const raw = planOrderedLikeRule0();
  const ordered = splitIntoActionWaves(raw).flatMap((w) => w.actions);

  it('reproduces the hazard in the RAW plan order', () => {
    // Guard the premise: without reordering, July's formula really does precede
    // July's creation. If this ever stops being true the test below is vacuous.
    const formulaIdx = raw.findIndex(
      (a) => a.type === 'SET_FORMULA' && String(a.formula).includes('July!'),
    );
    expect(formulaIdx).toBeGreaterThanOrEqual(0);
    expect(formulaIdx).toBeLessThan(createIndex(raw, 'July'));
  });

  it('creates every referenced sheet BEFORE the formula naming it', () => {
    for (const month of ['July', 'August', 'September', 'October', 'November', 'December']) {
      const created = createIndex(ordered, month);
      const usedAt = ordered.findIndex(
        (a) => a.type === 'SET_FORMULA' && String(a.formula ?? '').includes(`${month}!`),
      );
      expect(created).toBeGreaterThanOrEqual(0);
      expect(created).toBeLessThan(usedAt);
    }
  });

  it('loses and duplicates nothing while reordering', () => {
    expect(ordered).toHaveLength(raw.length);
    expect(new Set(ordered).size).toBe(raw.length);
  });

  it('puts every sheet creation ahead of every content and formula write', () => {
    const lastCreate = ordered.reduce(
      (acc, a, i) => (a.type === 'ADD_SHEET' || a.type === 'CREATE_SHEET' ? i : acc),
      -1,
    );
    const firstWrite = ordered.findIndex((a) =>
      ['SET_FORMULA', 'BATCH_SET', 'SET_CELL'].includes(String(a.type)),
    );
    expect(lastCreate).toBeLessThan(firstWrite);
  });

  it('keeps formatting after the content it formats', () => {
    const lastContent = ordered.reduce(
      (acc, a, i) => (['BATCH_SET', 'SET_CELL', 'SET_FORMULA'].includes(String(a.type)) ? i : acc),
      -1,
    );
    const firstFormat = ordered.findIndex((a) => a.type === 'FORMAT_RANGE');
    expect(firstFormat).toBeGreaterThan(lastContent);
  });
});
