import {
  applyPresentationPass,
  detectHeaderRuns,
  resolveCurrencyFormat,
} from '../src/excel-ai/utils/presentation-pass.util';
import { sheetsCreatedInBatch } from '../src/excel-ai/utils/sheet-header-state.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';
import { WorkbookContext } from '../src/types/cellix.types';

/**
 * TASKS.md #138 — deterministic presentation pass.
 *
 * The monthly-ledger build shipped 188 SET_CELL and one CREATE_CHART: correct
 * content, zero styling. These tests pin the styling this pass adds, and — just
 * as importantly — the two things it must never do: move content, or touch a
 * sheet the batch did not create.
 */

const HEADERS = [
  'Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out',
  'Nights', 'Rate Per Night', 'Total Amount', 'Source', 'Payment Status',
];

function monthSheetBatch(sheet = 'January'): SheetActionPayload[] {
  const actions: SheetActionPayload[] = [{ type: 'ADD_SHEET', name: sheet }];
  HEADERS.forEach((header, col) => {
    actions.push({ type: 'SET_CELL', sheetName: sheet, row: 3, col, value: header });
  });
  actions.push({ type: 'SET_CELL', sheetName: sheet, row: 1, col: 0, value: `${sheet} — Tracker` });
  return actions;
}

function findFormat(actions: SheetActionPayload[], row: number, col: number) {
  return actions.find(
    (a) => a.type === 'FORMAT_RANGE' && a.row === row && a.col === col,
  );
}

describe('presentation pass — header detection', () => {
  it('finds a header run on a sheet the batch creates', () => {
    const actions = monthSheetBatch();
    const runs = detectHeaderRuns(actions, sheetsCreatedInBatch(actions));
    expect(runs).toHaveLength(1);
    expect(runs[0].row).toBe(3);
    expect(runs[0].startCol).toBe(0);
    expect(runs[0].endCol).toBe(HEADERS.length - 1);
  });

  it('ignores sheets the batch did not create', () => {
    const actions: SheetActionPayload[] = [
      { type: 'SET_CELL', sheetName: 'Existing', row: 0, col: 0, value: 'A' },
      { type: 'SET_CELL', sheetName: 'Existing', row: 0, col: 1, value: 'B' },
      { type: 'SET_CELL', sheetName: 'Existing', row: 0, col: 2, value: 'C' },
    ];
    expect(detectHeaderRuns(actions, sheetsCreatedInBatch(actions))).toHaveLength(0);
  });

  it('needs at least three labels before a row counts as a header', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'Main' },
      { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'Dashboard' },
      { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 1, value: 'Totals' },
    ];
    expect(detectHeaderRuns(actions, sheetsCreatedInBatch(actions))).toHaveLength(0);
  });

  it('finds both header rows on a Main sheet with two tables', () => {
    const actions: SheetActionPayload[] = [{ type: 'ADD_SHEET', name: 'Main' }];
    ['Month', 'Total Amount', 'Paid Amount', 'Pending Amount'].forEach((h, col) =>
      actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 3, col, value: h }),
    );
    ['Month', 'Unit No', 'Guest', 'Total Amount'].forEach((h, col) =>
      actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 17, col, value: h }),
    );

    const runs = detectHeaderRuns(actions, sheetsCreatedInBatch(actions));
    expect(runs.map((r) => r.row)).toEqual([3, 17]);
  });
});

describe('presentation pass — emitted styling', () => {
  it('styles the header band', () => {
    const out = applyPresentationPass(monthSheetBatch());
    const header = findFormat(out, 3, 0);
    expect(header?.format).toMatchObject({
      bold: true,
      fillColor: '#2F5597',
      fontColor: '#FFFFFF',
      wrapText: true,
    });
    expect(header?.colCount).toBe(HEADERS.length);
  });

  it('freezes the header row and autofits the sheet', () => {
    const out = applyPresentationPass(monthSheetBatch());
    const freeze = out.find((a) => a.type === 'FREEZE_PANES');
    expect(freeze).toMatchObject({ sheetName: 'January', freezeRows: 4 });
    expect(out.some((a) => a.type === 'AUTOFIT_COLUMNS' && a.sheetName === 'January')).toBe(true);
  });

  it('number-formats money columns and count columns differently', () => {
    const out = applyPresentationPass(monthSheetBatch());
    const totalAmountCol = HEADERS.indexOf('Total Amount');
    const nightsCol = HEADERS.indexOf('Nights');

    expect(findFormat(out, 4, totalAmountCol)?.format?.numberFormat).toBe('#,##0.00');
    expect(findFormat(out, 4, nightsCol)?.format?.numberFormat).toBe('#,##0');
    // Text columns are left alone — 'Guest Name' contains the whole word
    // "Guest" and must still be vetoed as a name column (TASKS.md #158).
    expect(findFormat(out, 4, HEADERS.indexOf('Guest Name'))).toBeUndefined();
    expect(findFormat(out, 4, HEADERS.indexOf('Payment Status'))).toBeUndefined();
    // 'Unit No' is an identifier, not a quantity — 1001 must not become 1,001.
    expect(findFormat(out, 4, HEADERS.indexOf('Unit No'))).toBeUndefined();
  });

  it('styles a lone label above the header as a title', () => {
    const out = applyPresentationPass(monthSheetBatch());
    const title = findFormat(out, 1, 0);
    expect(title?.format).toMatchObject({ bold: true, fontSize: 14 });
  });

  it('appends only — every original action survives, in order', () => {
    const input = monthSheetBatch();
    const out = applyPresentationPass(input);
    expect(out.slice(0, input.length)).toEqual(input);
    expect(out.length).toBeGreaterThan(input.length);
  });

  it('never emits an action that moves or inserts content', () => {
    const input = monthSheetBatch();
    const added = applyPresentationPass(input).slice(input.length);
    const relocating = new Set(['INSERT_ROW', 'ADD_ROW', 'DELETE_ROW', 'INSERT_COLUMN', 'MOVE_RANGE']);
    expect(added.some((a) => relocating.has(a.type))).toBe(false);
  });

  it('is a no-op when the batch creates no sheets', () => {
    const input: SheetActionPayload[] = [
      { type: 'SET_CELL', sheetName: 'Existing', row: 0, col: 0, value: 'A' },
      { type: 'SET_CELL', sheetName: 'Existing', row: 0, col: 1, value: 'B' },
      { type: 'SET_CELL', sheetName: 'Existing', row: 0, col: 2, value: 'C' },
    ];
    expect(applyPresentationPass(input)).toEqual(input);
  });

  it('styles all twelve month sheets identically', () => {
    const actions = ['January', 'February', 'March'].flatMap((m) => monthSheetBatch(m));
    const out = applyPresentationPass(actions);
    for (const month of ['January', 'February', 'March']) {
      const header = out.find(
        (a) => a.type === 'FORMAT_RANGE' && a.sheetName === month && a.row === 3,
      );
      expect(header?.format?.fillColor).toBe('#2F5597');
      expect(out.some((a) => a.type === 'FREEZE_PANES' && a.sheetName === month)).toBe(true);
    }
  });
});

describe('presentation pass — currency is honoured, never invented', () => {
  it('defaults to a neutral thousands format', () => {
    expect(resolveCurrencyFormat({})).toBe('#,##0.00');
  });

  it('uses the currency the user named', () => {
    expect(resolveCurrencyFormat({ userMessage: 'track payments in INR' })).toBe('₹#,##0');
    expect(resolveCurrencyFormat({ userMessage: 'amounts in ₹' })).toBe('₹#,##0');
    expect(resolveCurrencyFormat({ userMessage: 'rate per night in dollars' })).toBe('$#,##0.00');
  });

  it('picks up a currency already used in the workbook', () => {
    const context = {
      activeSheet: 'Ledger',
      sheets: [
        {
          sheetName: 'Ledger',
          usedRange: 'A1:B2',
          rowCount: 2,
          colCount: 2,
          headers: ['Item', 'Amount'],
          sampleData: [],
          columnMeta: [{ numberFormat: '₹#,##0.00' }],
        },
      ],
    } as unknown as WorkbookContext;
    expect(resolveCurrencyFormat({ context })).toBe('₹#,##0');
  });
});

/**
 * TASKS.md #143 — March was silently left out of styling AND consolidation.
 *
 * A live smoke test of the monthly-ledger prompt showed the executor writing
 * eleven months' headers as ten SET_CELLs each, and ONE month's header (March)
 * as a single ADD_ROW carrying a `data` array of the same ten labels. Both are
 * legitimate ways to express "write this header row", but `detectHeaderRuns`
 * only recognized the first — so March got no header band, no FREEZE_PANES, no
 * AUTOFIT_COLUMNS, and (via `consolidation-pass.util.ts`, which reuses this
 * function) was silently missing from the consolidated VSTACK formula: typing
 * a booking into March would never have shown up on Main.
 */
describe('presentation pass — ADD_ROW-shaped headers (TASKS.md #143)', () => {
  function addRowHeaderBatch(sheet = 'March'): SheetActionPayload[] {
    return [
      { type: 'ADD_SHEET', name: sheet },
      { type: 'ADD_ROW', sheetName: sheet, data: [...HEADERS] },
    ];
  }

  it('detects an ADD_ROW header exactly like the equivalent SET_CELL header', () => {
    const setCellRuns = detectHeaderRuns(monthSheetBatch('January'), sheetsCreatedInBatch(monthSheetBatch('January')));
    const addRowRuns = detectHeaderRuns(addRowHeaderBatch('March'), sheetsCreatedInBatch(addRowHeaderBatch('March')));

    expect(addRowRuns).toHaveLength(1);
    expect(addRowRuns[0].row).toBe(0);
    expect(addRowRuns[0].startCol).toBe(0);
    expect(addRowRuns[0].endCol).toBe(HEADERS.length - 1);
    // Same schema, same shape — just a different action type carried it.
    expect([...addRowRuns[0].labels.values()]).toEqual([...setCellRuns[0].labels.values()]);
  });

  it('reads `values` as a fallback when `data` is absent', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'March' },
      { type: 'ADD_ROW', sheetName: 'March', values: [...HEADERS] } as SheetActionPayload,
    ];
    const runs = detectHeaderRuns(actions, sheetsCreatedInBatch(actions));
    expect(runs).toHaveLength(1);
    expect(runs[0].labels.size).toBe(HEADERS.length);
  });

  it('styles the ADD_ROW sheet identically to its SET_CELL siblings', () => {
    const mixed = [...monthSheetBatch('January'), ...addRowHeaderBatch('March')];
    const out = applyPresentationPass(mixed);

    for (const sheet of ['January', 'March']) {
      expect(out.some((a) => a.type === 'FREEZE_PANES' && a.sheetName === sheet)).toBe(true);
      expect(out.some((a) => a.type === 'AUTOFIT_COLUMNS' && a.sheetName === sheet)).toBe(true);
      const band = out.find(
        (a) => a.type === 'FORMAT_RANGE' && a.sheetName === sheet && a.format?.fillColor === '#2F5597',
      );
      expect(band?.colCount).toBe(HEADERS.length);
    }
  });

  it('never treats an ADD_ROW of actual data (mixed types, or on an existing sheet) as a header', () => {
    const dataRow: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'March' },
      { type: 'ADD_ROW', sheetName: 'March', data: ['A101', 2, 'John', '2026-03-01'] },
    ];
    expect(detectHeaderRuns(dataRow, sheetsCreatedInBatch(dataRow))).toHaveLength(0);

    const existingSheet: SheetActionPayload[] = [
      { type: 'ADD_ROW', sheetName: 'Existing', data: [...HEADERS] },
    ];
    expect(detectHeaderRuns(existingSheet, sheetsCreatedInBatch(existingSheet))).toHaveLength(0);
  });
});

/**
 * TASKS.md #144 — a sparse KPI row was mistaken for a table header.
 *
 * A live run's Main sheet wrote A1="Total Amount", C1="Paid Amount",
 * E1="Pending Amount" with B1/D1 left empty (their SUM formulas live one row
 * down). Three labels met MIN_HEADER_CELLS, and the run's inferred span (A:E)
 * was treated as a 5-wide header despite only 3 of those 5 columns actually
 * being labelled — painting a header band across two empty cells, landing
 * number formats on the label columns instead of the value columns beside
 * them, and stealing the sheet's one freeze/autofit slot from its real table
 * headers (row 3 and row 18 on the same sheet).
 */
describe('presentation pass — a header run must be contiguous (TASKS.md #144)', () => {
  function kpiRowPlusRealHeaders(): SheetActionPayload[] {
    const actions: SheetActionPayload[] = [{ type: 'ADD_SHEET', name: 'Main' }];
    // The sparse KPI row: labels at 0, 2, 4 — formula cells (not detected by
    // detectHeaderRuns, which only reads SET_CELL) sit at 1 and 3.
    actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'Total Amount' });
    actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 0, col: 2, value: 'Paid Amount' });
    actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 0, col: 4, value: 'Pending Amount' });
    // A genuine, contiguous table header lower on the same sheet.
    ['Month', 'Total Amount', 'Paid Amount', 'Pending Amount'].forEach((h, col) =>
      actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 3, col, value: h }),
    );
    return actions;
  }

  it('does not detect the sparse KPI row as a header run', () => {
    const actions = kpiRowPlusRealHeaders();
    const runs = detectHeaderRuns(actions, sheetsCreatedInBatch(actions));
    expect(runs.some((r) => r.row === 0)).toBe(false);
    expect(runs.map((r) => r.row)).toEqual([3]);
  });

  it('never paints a header band across the KPI row', () => {
    const out = applyPresentationPass(kpiRowPlusRealHeaders());
    expect(findFormat(out, 0, 0)).toBeUndefined();
  });

  it('gives the real table header the freeze/autofit that the KPI row would otherwise have stolen', () => {
    const out = applyPresentationPass(kpiRowPlusRealHeaders());
    const freeze = out.find((a) => a.type === 'FREEZE_PANES' && a.sheetName === 'Main');
    // freezeRows = header row (3) + 1, not the KPI row (0) + 1.
    expect(freeze?.freezeRows).toBe(4);
  });

  it('still detects a genuinely contiguous header even at row 0', () => {
    const actions: SheetActionPayload[] = [{ type: 'ADD_SHEET', name: 'Main' }];
    ['A', 'B', 'C'].forEach((h, col) =>
      actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 0, col, value: h }),
    );
    const runs = detectHeaderRuns(actions, sheetsCreatedInBatch(actions));
    expect(runs).toHaveLength(1);
  });
});

/**
 * TASKS.md #157 — BATCH_SET is the third shape a header arrives in.
 *
 * A live run wrote the Consolidated Transactions header as a BATCH_SET; the
 * detector understood only SET_CELL and ADD_ROW, so that table got no styling
 * AND `consolidation-pass.util.ts` (which reuses this detector) never found it.
 * #143 was this exact lesson for ADD_ROW — same fix, third shape.
 */
describe('presentation pass — BATCH_SET-shaped headers (TASKS.md #157)', () => {
  function batchSetHeader(sheet = 'Main', row = 18): SheetActionPayload[] {
    return [
      { type: 'ADD_SHEET', name: sheet },
      {
        type: 'BATCH_SET',
        sheetName: sheet,
        operations: HEADERS.map((h, i) => ({
          address: `${String.fromCharCode(65 + i)}${row}`,
          value: h,
        })),
      } as SheetActionPayload,
    ];
  }

  it('detects a header written as one BATCH_SET', () => {
    const actions = batchSetHeader();
    const runs = detectHeaderRuns(actions, sheetsCreatedInBatch(actions));
    expect(runs).toHaveLength(1);
    expect(runs[0].row).toBe(17);
    expect(runs[0].startCol).toBe(0);
    expect(runs[0].endCol).toBe(HEADERS.length - 1);
    expect([...runs[0].labels.values()]).toEqual(HEADERS);
  });

  it('styles it identically to its SET_CELL equivalent', () => {
    const out = applyPresentationPass(batchSetHeader());
    const band = out.find(
      (a) => a.type === 'FORMAT_RANGE' && a.format?.fillColor === '#2F5597',
    );
    expect(band?.row).toBe(17);
    expect(band?.colCount).toBe(HEADERS.length);
    expect(out.some((a) => a.type === 'FREEZE_PANES')).toBe(true);
  });

  it('ignores operations with no address or a numeric value', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'Main' },
      {
        type: 'BATCH_SET',
        sheetName: 'Main',
        operations: [
          { address: 'A1', value: 'Month' },
          { value: 'no address' },
          { address: 'C1', value: '42' },
        ],
      } as SheetActionPayload,
    ];
    // Only one usable label — below MIN_HEADER_CELLS, so no run.
    expect(detectHeaderRuns(actions, sheetsCreatedInBatch(actions))).toHaveLength(0);
  });
});
