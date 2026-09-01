import {
  applyConsolidationPass,
  buildConsolidationFormula,
  columnIndexToLetter,
  formulaSheetRef,
  planConsolidation,
} from '../src/excel-ai/utils/consolidation-pass.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #142 — "when I updated in the March sheet it won't update in Main".
 *
 * `Main`'s consolidated "All Bookings" table was written as a header row and
 * nothing else, so a booking typed into a month sheet never appeared there.
 * This pass adds one spilling dynamic-array formula that stacks every month
 * sheet, tags each row with its source, and drops the empty rows.
 */

const MONTHS = ['January', 'February', 'March'];

const MONTH_SCHEMA = [
  'Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out',
  'Rate Per Night', 'Total Amount', 'Source', 'Payment Status', 'Bank Account',
];

function ledgerBatch(): SheetActionPayload[] {
  const actions: SheetActionPayload[] = [];
  for (const month of MONTHS) actions.push({ type: 'ADD_SHEET', name: month });
  actions.push({ type: 'ADD_SHEET', name: 'Main' });

  for (const month of MONTHS) {
    MONTH_SCHEMA.forEach((h, col) =>
      actions.push({ type: 'SET_CELL', sheetName: month, row: 0, col, value: h }),
    );
  }

  // Main's own rollup table — starts with "Month" but is NOT a consolidation.
  ['Month', 'Total Amount', 'Paid Amount', 'Pending Amount'].forEach((h, col) =>
    actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 3, col, value: h }),
  );

  // Main's consolidated table — "Month" + the month sheets' exact schema.
  ['Month', ...MONTH_SCHEMA].forEach((h, col) =>
    actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 17, col, value: h }),
  );

  return actions;
}

describe('consolidation pass — helpers', () => {
  it('converts column indexes past Z', () => {
    expect(columnIndexToLetter(0)).toBe('A');
    expect(columnIndexToLetter(9)).toBe('J');
    expect(columnIndexToLetter(26)).toBe('AA');
  });

  it('quotes sheet names that need it, doubling embedded quotes', () => {
    expect(formulaSheetRef('January')).toBe('January');
    expect(formulaSheetRef('Jan 2026')).toBe("'Jan 2026'");
    expect(formulaSheetRef("Bob's")).toBe("'Bob''s'");
  });
});

describe('consolidation pass — plan detection', () => {
  it('targets the table whose header is "Month" + the month sheets\' schema', () => {
    const plan = planConsolidation(ledgerBatch());
    expect(plan).not.toBeNull();
    expect(plan!.targetSheet).toBe('Main');
    // Written directly under the header row (17), at its first column.
    expect(plan!.row).toBe(18);
    expect(plan!.col).toBe(0);
  });

  it('does not mistake the Monthly Totals rollup for a consolidation', () => {
    // Both headers begin with "Month"; only exact schema equality separates them.
    const plan = planConsolidation(ledgerBatch());
    expect(plan!.row).not.toBe(4);
  });

  it('lists source sheets in creation order, not alphabetical', () => {
    const plan = planConsolidation(ledgerBatch());
    expect(plan!.sourceSheets).toEqual(['January', 'February', 'March']);
  });

  it('excludes the target sheet from its own sources', () => {
    const plan = planConsolidation(ledgerBatch());
    expect(plan!.sourceSheets).not.toContain('Main');
  });

  it('needs at least two source sheets', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'January' },
      { type: 'ADD_SHEET', name: 'Main' },
    ];
    MONTH_SCHEMA.forEach((h, col) =>
      actions.push({ type: 'SET_CELL', sheetName: 'January', row: 0, col, value: h }),
    );
    ['Month', ...MONTH_SCHEMA].forEach((h, col) =>
      actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 17, col, value: h }),
    );
    expect(planConsolidation(actions)).toBeNull();
  });

  it('is a no-op when no table consolidates anything', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'Data' },
      { type: 'SET_CELL', sheetName: 'Data', row: 0, col: 0, value: 'A' },
      { type: 'SET_CELL', sheetName: 'Data', row: 0, col: 1, value: 'B' },
      { type: 'SET_CELL', sheetName: 'Data', row: 0, col: 2, value: 'C' },
    ];
    expect(applyConsolidationPass(actions)).toEqual(actions);
  });
});

describe('consolidation pass — the formula', () => {
  const formula = buildConsolidationFormula(['January', 'February'], 0, 9);

  it('stacks every source sheet over its data range, below the header', () => {
    expect(formula).toContain('January!A2:J500');
    expect(formula).toContain('February!A2:J500');
    expect(formula.startsWith('=LET(rows,VSTACK(')).toBe(true);
  });

  it('tags each row with the sheet it came from', () => {
    expect(formula).toContain('EXPAND("January"');
    expect(formula).toContain('EXPAND("February"');
  });

  it('keeps a row when ANY data column is filled, not just the first', () => {
    // The reported case: Guest and Guest Name filled, Unit No left empty. A
    // first-column test would have hidden that booking.
    expect(formula).toContain('BYROW(DROP(rows,,1),LAMBDA(r,COUNTA(r)>0))');
  });

  it('filters empty rows out and degrades to blank rather than an error', () => {
    expect(formula).toContain('IFERROR(FILTER(rows,keep,""),"")');
  });

  it('quotes sheet names with spaces', () => {
    expect(buildConsolidationFormula(['Jan 2026'], 0, 9)).toContain("'Jan 2026'!A2:J500");
  });
});

describe('consolidation pass — emitted action', () => {
  it('appends exactly one SET_FORMULA, leaving the batch otherwise untouched', () => {
    const input = ledgerBatch();
    const out = applyConsolidationPass(input, { dynamicArrays: true });
    expect(out.slice(0, input.length)).toEqual(input);
    expect(out).toHaveLength(input.length + 1);

    const added = out[out.length - 1];
    expect(added.type).toBe('SET_FORMULA');
    expect(added.sheetName).toBe('Main');
    expect(added.row).toBe(18);
    expect(added.col).toBe(0);
    expect(added.formula).toContain('VSTACK');
  });
});

/**
 * TASKS.md #143 — the same ADD_ROW-shaped header gap, in consolidation.
 *
 * `planConsolidation` calls `detectHeaderRuns` to find both the target table
 * and its source sheets. Before the #143 fix, a month sheet whose header
 * arrived as ADD_ROW (rather than SET_CELL) was invisible to it — silently
 * missing from the VSTACK, so a booking typed into that one sheet would never
 * reach Main's consolidated view while its siblings worked correctly.
 */
describe('consolidation pass — ADD_ROW-shaped source sheets', () => {
  function mixedShapeBatch(): SheetActionPayload[] {
    const actions: SheetActionPayload[] = [];
    for (const month of MONTHS) actions.push({ type: 'ADD_SHEET', name: month });
    actions.push({ type: 'ADD_SHEET', name: 'Main' });

    // January, February: header as SET_CELLs (the common shape).
    for (const month of ['January', 'February']) {
      MONTH_SCHEMA.forEach((h, col) =>
        actions.push({ type: 'SET_CELL', sheetName: month, row: 0, col, value: h }),
      );
    }
    // March: header as ONE ADD_ROW — the shape a live run actually produced.
    actions.push({ type: 'ADD_ROW', sheetName: 'March', data: [...MONTH_SCHEMA] });

    ['Month', ...MONTH_SCHEMA].forEach((h, col) =>
      actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 17, col, value: h }),
    );
    return actions;
  }

  it('includes the ADD_ROW-shaped sheet in the source list', () => {
    const plan = planConsolidation(mixedShapeBatch());
    expect(plan).not.toBeNull();
    expect(plan!.sourceSheets).toEqual(expect.arrayContaining(['January', 'February', 'March']));
  });

  it('puts March in the actual VSTACK formula, not just the plan', () => {
    const out = applyConsolidationPass(mixedShapeBatch(), { dynamicArrays: true });
    const formula = out[out.length - 1].formula!;
    expect(formula).toContain('"March"');
    expect(formula).toContain('March!A2:J500');
  });
});

/**
 * TASKS.md #152 — the consolidation formula family is chosen against a probed
 * fact, not an assumption. On a host without dynamic arrays the spilling form
 * is a single `#NAME?` cell and the whole consolidated view is silently empty.
 */
describe('consolidation pass — host capability decides the formula family', () => {
  it('uses the spilling dynamic-array form when support is PROVEN', () => {
    const out = applyConsolidationPass(ledgerBatch(), { dynamicArrays: true });
    const added = out.slice(ledgerBatch().length);
    expect(added).toHaveLength(1);
    expect(added[0].formula).toContain('VSTACK');
  });

  it('writes a single explanatory note when the host lacks dynamic arrays', () => {
    const out = applyConsolidationPass(ledgerBatch(), { dynamicArrays: false });
    const added = out.slice(ledgerBatch().length);
    // One cell, not a thousand formulas — see LEGACY_NOTE for why two per-cell
    // fallbacks were built, measured and rejected.
    expect(added).toHaveLength(1);
    expect(added[0].type).toBe('SET_CELL');
    expect(String(added[0].value)).toContain('Excel 365');
    expect(added.some((a) => a.formula?.includes('VSTACK'))).toBe(false);
  });

  it('treats UNPROBED as unsupported - a wrong guess is a silent #NAME?', () => {
    const out = applyConsolidationPass(ledgerBatch());
    const added = out.slice(ledgerBatch().length);
    expect(added.some((a) => a.formula?.includes('VSTACK'))).toBe(false);
  });

  it('the note lands directly under the consolidated header', () => {
    const out = applyConsolidationPass(ledgerBatch(), { dynamicArrays: false });
    const added = out.slice(ledgerBatch().length);
    expect(added[0].row).toBe(18);
    expect(added[0].col).toBe(0);
    expect(added[0].sheetName).toBe('Main');
  });
});
