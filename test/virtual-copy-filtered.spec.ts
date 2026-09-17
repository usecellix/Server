import { buildShadowWorkbook, shadowSheetToContext } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import { WorkbookContext } from '../src/agents/types/agent.types';

describe('virtualApply COPY_FILTERED_RANGE / MOVE_RANGE', () => {
  const context: WorkbookContext = {
    activeSheetName: 'Purchase Register',
    sheets: [
      {
        name: 'Purchase Register',
        usedRange: 'A1:B4',
        rowCount: 4,
        columnCount: 2,
        values: [
          ['Vendor', 'Payment Status'],
          ['Acme', 'Pending'],
          ['Beta', 'Paid'],
          ['Gamma', 'Pending'],
        ],
        formulas: [[], [], [], []],
        numberFormats: [[], [], [], []],
        structure: 'data_table',
      headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('COPY_FILTERED_RANGE is deliberately NOT simulated — reads the SOURCE through the same shadow that mispredicted a live full-sheet copy (only ~11 of 61 rows known), so it is no longer previewed (reversibility-catalog.ts marks it non-reversible)', () => {
    const shadow = buildShadowWorkbook(context);
    const after = virtualApply(shadow, [
      {
        type: 'ADD_SHEET',
        name: 'Pending Payments',
      },
      {
        type: 'COPY_FILTERED_RANGE',
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:B4',
        hasHeaders: true,
        destSheet: 'Pending Payments',
        destStartCell: 'A1',
        filter: {
          column: 'Payment Status',
          operator: 'equals',
          value: 'Pending',
        },
        mode: 'copy',
      },
    ]);

    // ADD_SHEET still runs (a different action type) but the copy itself is a no-op.
    const dest = shadowSheetToContext(after.sheets.get('Pending Payments')!);
    expect(dest.values.length).toBe(0);

    // Source is completely untouched — Office.js is the sole source of truth.
    const source = shadowSheetToContext(after.sheets.get('Purchase Register')!);
    expect(source.values[1]).toEqual(['Acme', 'Pending']);
    expect(source.values[2]).toEqual(['Beta', 'Paid']);
  });

  it('MOVE_RANGE relocates an entire block and clears the source', () => {
    const shadow = buildShadowWorkbook(context);
    const after = virtualApply(shadow, [
      {
        type: 'MOVE_RANGE',
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:B2',
        destSheet: 'Archive',
        destStartCell: 'A1',
      },
    ]);

    const dest = shadowSheetToContext(after.sheets.get('Archive')!);
    expect(dest.values[0]).toEqual(['Vendor', 'Payment Status']);
    expect(dest.values[1]).toEqual(['Acme', 'Pending']);

    const source = shadowSheetToContext(after.sheets.get('Purchase Register')!);
    expect(source.values[0]?.[0]).toBeNull();
    expect(source.values[1]?.[0]).toBeNull();
  });
});
