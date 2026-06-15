import { normalizeSortRangeAddress, stripSheetPrefix } from '../src/agents/utils/range-address.util';
import { compareSortValues, parseSortableValue } from '../src/agents/utils/sort-value.util';
import { buildSortFallbackAction } from '../src/agents/utils/sort-action.util';
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';

describe('range-address.util', () => {
  it('strips sheet prefix from used range addresses', () => {
    expect(stripSheetPrefix("'Purchases'!A1:M339")).toBe('A1:M339');
    expect(stripSheetPrefix('Sheet1!A1:F50')).toBe('A1:F50');
  });

  it('normalizes sort range to local A1 notation', () => {
    expect(
      normalizeSortRangeAddress({
        usedRange: "'Purchases'!A1:M339",
        rowCount: 339,
        columnCount: 13,
        key: 5,
      }),
    ).toBe('A1:M339');
  });
});

describe('sort-value.util accounting values', () => {
  it('parses CGST amounts with Dr suffix', () => {
    expect(parseSortableValue('8533.98 Dr')).toBe(8533.98);
    expect(parseSortableValue('450.09 Dr')).toBe(450.09);
  });

  it('sorts accounting values numerically', () => {
    const values = ['8533.98 Dr', '450.09 Dr', '12088.13 Dr'];
    const sorted = [...values].sort(compareSortValues);
    expect(sorted).toEqual(['450.09 Dr', '8533.98 Dr', '12088.13 Dr']);
  });
});

describe('buildSortFallbackAction range output', () => {
  it('emits local range without sheet prefix', () => {
    const context: WorkbookContext = {
      activeSheetName: 'Purchases',
      sheets: [
        {
          name: 'CGST Sorted',
          usedRange: "'Purchases'!A1:M339",
          rowCount: 339,
          columnCount: 13,
          values: [
            ['Date', 'Particulars', 'GSTIN/UIN', 'Purchase@0%', 'Purchase@5%', 'CGST', 'SGST'],
            ['20-Feb-25', 'EM Roof', '', '', '', '8533.98 Dr', '8533.98 Dr'],
          ],
          formulas: [],
          numberFormats: [],
          structure: 'data_table',
        },
      ],
      namedRanges: [],
      tables: [],
    };

    const subtask: SubTask = {
      id: 's2',
      description: 'sort the values of CGST in ascending order on sheet "CGST Sorted"',
      targetSheet: 'CGST Sorted',
      dependsOn: ['s1'],
      estimatedActions: 1,
    };

    const action = buildSortFallbackAction(subtask, context);
    expect(action?.range).toBe('A1:M339');
    expect(action?.sheetName).toBe('CGST Sorted');
    expect(action?.key).toBe(5);
  });

  it('extracts Date from "sort the sheet based on the Date"', () => {
    const context: WorkbookContext = {
      activeSheetName: 'Purchases',
      sheets: [
        {
          name: 'Purchases',
          usedRange: 'A1:M339',
          rowCount: 339,
          columnCount: 13,
          values: [
            ['Date', 'Particulars', 'GSTIN/UIN', 'Purchase@0%', 'Purchase@5%', 'CGST', 'SGST'],
            ['20-Feb-25', 'EM Roof', '', '', '', '8533.98 Dr', '8533.98 Dr'],
            ['05-Jan-25', 'Fabs Trading', '', '', '', '450.09 Dr', '450.09 Dr'],
          ],
          formulas: [],
          numberFormats: [],
          structure: 'data_table',
        },
      ],
      namedRanges: [],
      tables: [],
    };

    const subtask: SubTask = {
      id: 's1',
      description: 'sort the sheet based on the Date',
      targetSheet: 'Purchases',
      dependsOn: [],
      estimatedActions: 1,
    };

    const action = buildSortFallbackAction(subtask, context);
    expect(action).not.toBeNull();
    expect(action?.columnName).toBe('Date');
    expect(action?.key).toBe(0);
    expect(action?.ascending).toBe(true);
  });
});
