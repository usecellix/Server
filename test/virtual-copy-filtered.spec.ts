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
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('copies filtered rows into a destination sheet', () => {
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

    const dest = shadowSheetToContext(after.sheets.get('Pending Payments')!);
    expect(dest.values[0]).toEqual(['Vendor', 'Payment Status']);
    expect(dest.values[1]).toEqual(['Acme', 'Pending']);
    expect(dest.values[2]).toEqual(['Gamma', 'Pending']);
    expect(dest.values.length).toBe(3);

    const source = shadowSheetToContext(after.sheets.get('Purchase Register')!);
    expect(source.values[1]).toEqual(['Acme', 'Pending']);
    expect(source.values[2]).toEqual(['Beta', 'Paid']);
  });

  it('move mode clears matched source rows after copy', () => {
    const shadow = buildShadowWorkbook(context);
    const after = virtualApply(shadow, [
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
        mode: 'move',
      },
    ]);

    const dest = shadowSheetToContext(after.sheets.get('Pending Payments')!);
    expect(dest.values.length).toBe(3);

    const source = shadowSheetToContext(after.sheets.get('Purchase Register')!);
    expect(source.values[1]?.[0]).toBeNull();
    expect(source.values[1]?.[1]).toBeNull();
    expect(source.values[2]).toEqual(['Beta', 'Paid']);
    expect(source.values[3]?.[0]).toBeNull();
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
