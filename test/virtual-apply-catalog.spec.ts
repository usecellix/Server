import { WorkbookContext } from '../src/agents/types/agent.types';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import { generateDiff } from '../src/audit/diff.engine';
import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import {
  VIRTUAL_APPLY_CATALOG,
  ALL_VIRTUAL_APPLY_TYPES,
  UNSIMULATED_ACTION_TYPES,
} from '../src/virtual/virtual-apply-catalog';

describe('virtual-apply-catalog (TASKS.md #41)', () => {
  it('gives every documented no-op a real, non-empty reason', () => {
    for (const type of UNSIMULATED_ACTION_TYPES) {
      const entry = VIRTUAL_APPLY_CATALOG[type];
      expect(entry.simulated).toBe(false);
      if (!entry.simulated) {
        expect(entry.reason.trim().length).toBeGreaterThan(10);
      }
    }
  });

  it('marks the four gaps found during the #41 audit as simulated, now that #66 fixed them', () => {
    const fixed = ['CLEAR_CONTENT', 'CLEAR_ALL', 'SET_MATCHING_ROWS', 'MERGE_CELLS'];
    for (const type of fixed) {
      expect(VIRTUAL_APPLY_CATALOG[type as keyof typeof VIRTUAL_APPLY_CATALOG].simulated).toBe(true);
    }
  });

  it('has exactly the action types the backend catalog declares — no more, no fewer', () => {
    // Record<SheetActionType, X> already makes this a compile error if it
    // drifts; asserting it here too documents the intent at runtime.
    expect(ALL_VIRTUAL_APPLY_TYPES.length).toBe(Object.keys(VIRTUAL_APPLY_CATALOG).length);
  });
});

describe('virtualApply — DELETE_SHEET (TASKS.md #41 gap fix)', () => {
  it('actually removes the sheet from the shadow workbook', () => {
    const context: WorkbookContext = {
      activeSheetName: 'Keep',
      sheets: [
        { name: 'Keep', usedRange: 'A1', rowCount: 1, columnCount: 1, values: [['x']], formulas: [['']], numberFormats: [['General']], structure: 'data_table', headerRowIndex: 0 },
        { name: 'DropMe', usedRange: 'A1', rowCount: 1, columnCount: 1, values: [['y']], formulas: [['']], numberFormats: [['General']], structure: 'data_table', headerRowIndex: 0 },
      ],
      namedRanges: [],
      tables: [],
    };
    const before = buildShadowWorkbook(context);
    expect(before.sheets.has('DropMe')).toBe(true);

    const after = virtualApply(before, [
      { type: 'DELETE_SHEET', sheetName: 'DropMe' },
    ] as never);

    expect(after.sheets.has('DropMe')).toBe(false);
    expect(after.sheets.has('Keep')).toBe(true);
  });
});

describe('virtualApply — FILL_DOWN / FILL_RIGHT simulation (TASKS.md #42)', () => {
  const baseContext: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      {
        name: 'Sheet1',
        usedRange: 'A1:C1',
        rowCount: 5,
        columnCount: 3,
        values: [['10', '20', '30']],
        formulas: [['', '=B1*2', '']],
        numberFormats: [['General', 'General', 'General']],
        structure: 'data_table',
        headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('shifts a relative row reference correctly when filling down (sourceRange/targetRange shape)', () => {
    const before = buildShadowWorkbook(baseContext);
    const after = virtualApply(before, [
      { type: 'FILL_DOWN', sheetName: 'Sheet1', sourceRange: 'B1', targetRange: 'B2:B3' },
    ] as never);

    expect(after.sheets.get('Sheet1')?.cells.get('B2')?.formula).toBe('=B2*2');
    expect(after.sheets.get('Sheet1')?.cells.get('B3')?.formula).toBe('=B3*2');
  });

  it('shifts a relative row reference correctly when filling down (row/col/endRow shape — what the Executor actually emits)', () => {
    // {"type":"FILL_DOWN","col":3,"row":1,"endRow":847} is the Executor's own
    // documented prompt example (cellix-system-prompt.ts) — this shape, not
    // sourceRange/targetRange, is what virtualApply actually receives from
    // real traffic, since normalize-executor-output.util.ts never renames
    // FILL_DOWN's fields before virtualApply sees the action.
    const before = buildShadowWorkbook(baseContext);
    const after = virtualApply(before, [
      { type: 'FILL_DOWN', sheetName: 'Sheet1', row: 0, col: 1, endRow: 2 },
    ] as never);

    // Source B1 = "=B1*2" shifted down 1 row -> "=B2*2", down 2 rows -> "=B3*2".
    expect(after.sheets.get('Sheet1')?.cells.get('B2')?.formula).toBe('=B2*2');
    expect(after.sheets.get('Sheet1')?.cells.get('B3')?.formula).toBe('=B3*2');
  });

  it('shifts a relative column reference correctly when filling right', () => {
    const before = buildShadowWorkbook(baseContext);
    const after = virtualApply(before, [
      { type: 'FILL_RIGHT', sheetName: 'Sheet1', row: 0, col: 1, endCol: 2 },
    ] as never);

    // Source B1 = "=B1*2" shifted right 1 col -> "=C1*2".
    expect(after.sheets.get('Sheet1')?.cells.get('C1')?.formula).toBe('=C1*2');
  });

  it('catches a fill-right reference shifted out of the sheet\'s bounds via checkPostApply — before any real Excel write', () => {
    // 3-column sheet (A,B,C). Source B1 references the sheet's LAST column
    // (C1). Filling right into C1 shifts that reference to D1, which does
    // not exist (columnCount = 3) — the exact "formula error introduced by
    // this change" class checkPostApply exists to catch.
    const context: WorkbookContext = {
      activeSheetName: 'Sheet1',
      sheets: [
        {
          name: 'Sheet1',
          usedRange: 'A1:C1',
          rowCount: 1,
          columnCount: 3,
          values: [['1', '2', '3']],
          formulas: [['', '=C1*2', '']],
          numberFormats: [['General', 'General', 'General']],
          structure: 'data_table',
          headerRowIndex: 0,
        },
      ],
      namedRanges: [],
      tables: [],
    };

    const before = buildShadowWorkbook(context);
    const fillRight = { type: 'FILL_RIGHT', sheetName: 'Sheet1', row: 0, col: 1, endCol: 2 };
    const after = virtualApply(before, [fillRight] as never);

    // Confirm the shadow actually holds the bad reference (not silently
    // skipped, per the pre-#42 no-op behavior).
    expect(after.sheets.get('Sheet1')?.cells.get('C1')?.formula).toBe('=D1*2');

    const validator = new FormulaValidatorService();
    const result = validator.checkPostApply(after, [fillRight] as never, context, 'Sheet1');

    expect(result.passed).toBe(false);
    expect(result.issues.some((i) => /outside sheet.*bounds/i.test(i.message))).toBe(true);
  });

  it('does nothing when sourceRange/targetRange cannot be resolved, rather than throwing', () => {
    const before = buildShadowWorkbook(baseContext);
    expect(() =>
      virtualApply(before, [
        { type: 'FILL_DOWN', sheetName: 'Sheet1', sourceRange: '', targetRange: '' },
      ] as never),
    ).not.toThrow();
  });
});

describe('virtualApply — CLEAR_CONTENT / CLEAR_ALL / SET_MATCHING_ROWS / MERGE_CELLS (TASKS.md #66)', () => {
  const purchaseContext: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      {
        name: 'Sheet1',
        usedRange: 'A1:B3',
        rowCount: 3,
        columnCount: 2,
        values: [
          ['Item', 'Status'],
          ['Apple', 'Pending'],
          ['Pear', 'Pending'],
        ],
        formulas: [['', ''], ['', ''], ['', '']],
        numberFormats: [['General', 'General'], ['General', 'General'], ['General', 'General']],
        structure: 'data_table',
        headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('CLEAR_CONTENT clears the value but keeps the number format', () => {
    const before = buildShadowWorkbook(purchaseContext);
    const after = virtualApply(before, [
      { type: 'CLEAR_CONTENT', sheetName: 'Sheet1', row: 1, col: 0, rowCount: 1, colCount: 1 },
    ] as never);

    const cell = after.sheets.get('Sheet1')?.cells.get('A2');
    expect(cell?.value).toBeNull();
    expect(cell?.numberFormat).toBe('General');
  });

  it('CLEAR_ALL clears both value and resets the number format', () => {
    const before = buildShadowWorkbook(purchaseContext);
    // Give A2 a distinct format first so we can prove CLEAR_ALL actually resets it.
    const formatted = virtualApply(before, [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 0, value: 'Apple', format: { numberFormat: 'dd/mm/yyyy' } },
    ] as never);
    expect(formatted.sheets.get('Sheet1')?.cells.get('A2')?.numberFormat).toBe('dd/mm/yyyy');

    const after = virtualApply(formatted, [
      { type: 'CLEAR_ALL', sheetName: 'Sheet1', row: 1, col: 0, rowCount: 1, colCount: 1 },
    ] as never);
    const cell = after.sheets.get('Sheet1')?.cells.get('A2');
    expect(cell?.value).toBeNull();
    expect(cell?.numberFormat).toBe('General');
  });

  it('SET_MATCHING_ROWS writes the target value into every row matching the filter', () => {
    const before = buildShadowWorkbook(purchaseContext);
    const after = virtualApply(before, [
      {
        type: 'SET_MATCHING_ROWS',
        sheetName: 'Sheet1',
        range: 'A1:B3',
        hasHeaders: true,
        filter: { column: 'Item', operator: 'equals', value: 'Apple' },
        targetColumn: 'Status',
        value: 'Paid',
      },
    ] as never);

    expect(after.sheets.get('Sheet1')?.cells.get('B2')?.value).toBe('Paid');
    // Pear did not match the filter — untouched.
    expect(after.sheets.get('Sheet1')?.cells.get('B3')?.value).toBe('Pending');
  });

  it('SET_MATCHING_ROWS with no filter writes every data row', () => {
    const before = buildShadowWorkbook(purchaseContext);
    const after = virtualApply(before, [
      {
        type: 'SET_MATCHING_ROWS',
        sheetName: 'Sheet1',
        range: 'A1:B3',
        hasHeaders: true,
        targetColumn: 'Status',
        value: 'Reviewed',
      },
    ] as never);

    expect(after.sheets.get('Sheet1')?.cells.get('B2')?.value).toBe('Reviewed');
    expect(after.sheets.get('Sheet1')?.cells.get('B3')?.value).toBe('Reviewed');
  });

  it('MERGE_CELLS keeps only the top-left value and discards the rest — and the diff shows it', () => {
    const before = buildShadowWorkbook(purchaseContext);
    const after = virtualApply(before, [
      { type: 'MERGE_CELLS', sheetName: 'Sheet1', row: 0, col: 0, rowCount: 1, colCount: 2 },
    ] as never);

    expect(after.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item'); // top-left survives
    expect(after.sheets.get('Sheet1')?.cells.get('B1')?.value).toBeNull(); // discarded, exactly like real Excel

    // The generic diff mechanism the preview UI already relies on picks up
    // the discarded value as a real change — no new checker needed, per
    // TASKS.md #66's note that this is the natural payoff of accurate
    // simulation, not a separate feature.
    const changes = generateDiff(before, after);
    const discarded = changes.find((c) => c.cell === 'B1');
    expect(discarded).toBeDefined();
    expect(discarded?.before).toBe('Status');
    expect(discarded?.after).toBeNull();
  });

  it('MERGE_CELLS on a single cell is a no-op', () => {
    const before = buildShadowWorkbook(purchaseContext);
    const after = virtualApply(before, [
      { type: 'MERGE_CELLS', sheetName: 'Sheet1', row: 0, col: 0, rowCount: 1, colCount: 1 },
    ] as never);
    expect(after.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item');
  });
});
