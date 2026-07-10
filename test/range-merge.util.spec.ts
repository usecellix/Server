import { mergeRangeIntoSheet, parseA1Range } from '../src/agents/utils/range-merge.util';
import { SheetContext } from '../src/agents/types/agent.types';

describe('range-merge.util', () => {
  it('parses A1 ranges', () => {
    expect(parseA1Range('A1:C10')).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 9,
      endCol: 2,
    });
  });

  it('merges fetched values into sparse sheet context', () => {
    const sheet: SheetContext = {
      name: 'Sheet1',
      usedRange: 'A1:C1000',
      rowCount: 1000,
      columnCount: 3,
      values: [['Name', 'Qty', 'Price']],
      formulas: [['', '', '']],
      numberFormats: [['General', 'General', 'General']],
      structure: 'data_table',
      dataTruncated: true,
    };

    const merged = mergeRangeIntoSheet(sheet, 'A2:C4', [
      ['Apple', 10, 1.5],
      ['Banana', 5, 0.75],
      ['Cherry', 2, 2],
    ]);

    expect(merged.values[1]).toEqual(['Apple', 10, 1.5]);
    expect(merged.values[3]).toEqual(['Cherry', 2, 2]);
    expect(merged.rowCount).toBe(1000);
  });
});
