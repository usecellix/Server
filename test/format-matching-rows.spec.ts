import {
  findMatchingRowOffsets,
  resolveFilterColumnIndex,
} from '../src/agents/utils/range-filter.util';
import {
  buildHeaderRowFormatAction,
  isClearHighlightMessage,
  isHeaderRowFormatRequest,
  normalizeFormatMatchingRowsAction,
  normalizeTier1ConditionalFormatActions,
  normalizeTier1HeaderFormatActions,
} from '../src/excel-ai/utils/format-matching-rows.util';
import { WorkbookContext } from '../src/agents/types/agent.types';
import { SheetAction } from '../src/excel-ai/types/sheet-actions.types';

describe('FORMAT_MATCHING_ROWS helpers', () => {
  const headers = [
    'Date',
    'Invoice No',
    'Supplier',
    'GSTIN',
    'Item',
    'Qty',
    'Unit Price',
    'Tax %',
    'Tax Amount',
    'Total Amount',
    'Payment Status',
    'Remarks',
  ];

  const rows = [
    headers,
    ['2026-01-01', 'INV-1', 'A', 'G1', 'Item', 1, 10, 18, 1.8, 11.8, 'Paid', ''],
    ['2026-01-02', 'INV-2', 'B', 'G2', 'Item', 1, 10, 18, 1.8, 11.8, 'Pending', ''],
    ['2026-01-03', 'INV-3', 'C', 'G3', 'Item', 1, 10, 18, 1.8, 11.8, 'Pending', ''],
    ['2026-01-04', 'INV-4', 'D', 'G4', 'Item', 1, 10, 18, 1.8, 11.8, 'Paid', ''],
  ];

  const workbookContext: WorkbookContext = {
    activeSheetName: 'Purchase Register',
    sheets: [
      {
        name: 'Purchase Register',
        usedRange: "'Purchase Register'!A1:L5",
        rowCount: 5,
        columnCount: 12,
        values: rows,
        formulas: [],
        numberFormats: [],
        structure: 'data_table',
      headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('resolves 1-based Excel column 11 to Payment Status index', () => {
    expect(resolveFilterColumnIndex(headers, 11)).toBe(10);
  });

  it('finds Pending row offsets within the range', () => {
    expect(
      findMatchingRowOffsets(rows, true, {
        column: 'Payment Status',
        operator: 'equals',
        value: 'Pending',
      }),
    ).toEqual([2, 3]);
  });

  it('rewrites HIGHLIGHT_CELL + condition into FORMAT_MATCHING_ROWS', () => {
    const broken = {
      type: 'HIGHLIGHT_CELL',
      sheetName: 'Purchase Register',
      row: 1,
      col: 0,
      rowCount: 55,
      colCount: 13,
      condition: { type: 'TEXT_EQ', column: 11, value: 'Pending' },
      format: { fillColor: '#FFC7CE' },
    } as SheetAction;

    const normalized = normalizeFormatMatchingRowsAction(broken, workbookContext);
    expect(normalized).toEqual({
      type: 'FORMAT_MATCHING_ROWS',
      sheetName: 'Purchase Register',
      range: 'A1:L5',
      hasHeaders: true,
      filter: {
        column: 'Payment Status',
        operator: 'equals',
        value: 'Pending',
      },
      format: { fillColor: '#FFC7CE' },
    });
  });

  it('normalizes a clean FORMAT_MATCHING_ROWS action', () => {
    const action: SheetAction = {
      type: 'FORMAT_MATCHING_ROWS',
      sheetName: 'Purchase Register',
      range: "'Purchase Register'!A1:L51",
      hasHeaders: true,
      filter: {
        column: 'Payment Status',
        operator: 'equals',
        value: 'Pending',
      },
      format: { fillColor: '#FFC7CE' },
    };

    const normalized = normalizeFormatMatchingRowsAction(action, workbookContext);
    expect(normalized.type).toBe('FORMAT_MATCHING_ROWS');
    expect(normalized.range).toBe('A1:L51');
    expect(normalized.filter?.column).toBe('Payment Status');
  });

  it('detects clear-highlight phrasing', () => {
    expect(isClearHighlightMessage('remvoe the highlights red')).toBe(true);
    expect(isClearHighlightMessage('clear the red fill from pending rows')).toBe(true);
  });

  it('applies clearFill when removing highlights', () => {
    const action: SheetAction = {
      type: 'FORMAT_MATCHING_ROWS',
      sheetName: 'Purchase Register',
      range: 'A1:L51',
      hasHeaders: true,
      filter: {
        column: 'Payment Status',
        operator: 'equals',
        value: 'Pending',
      },
      format: { fillColor: '#FFFFFF' },
    };

    const normalized = normalizeFormatMatchingRowsAction(
      action,
      workbookContext,
      'remove the red highlights',
    );
    expect(normalized.format).toEqual({ clearFill: true });
  });

  describe('Spec 24 header row FORMAT_RANGE', () => {
    it('detects header-row format requests', () => {
      expect(isHeaderRowFormatRequest('Highlight the header row with light bg green')).toBe(true);
      expect(isHeaderRowFormatRequest('highlight pending rows light green')).toBe(false);
    });

    it('rewrites misrouted FORMAT_MATCHING_ROWS to FORMAT_RANGE for header phrasing', () => {
      const bad: SheetAction = {
        type: 'FORMAT_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A1:L5',
        hasHeaders: false,
        filter: { column: 'Date', operator: 'equals', value: 'Date' },
        format: { fillColor: '#C6EFCE' },
      };

      const out = normalizeTier1ConditionalFormatActions(
        [bad],
        workbookContext,
        'Highlight the header row with light bg green',
      );
      expect(out).toHaveLength(1);
      expect(out[0]).toMatchObject({
        type: 'FORMAT_RANGE',
        sheetName: 'Purchase Register',
        row: 0,
        col: 0,
        rowCount: 1,
        format: { fillColor: '#C6EFCE' },
      });
      expect(out[0]!.colCount).toBeGreaterThanOrEqual(12);
    });

    it('buildHeaderRowFormatAction / HEADER_FORMAT hint yields one FORMAT_RANGE', () => {
      const action = buildHeaderRowFormatAction(
        workbookContext,
        'Highlight the header row with light bg green',
      );
      expect(action.type).toBe('FORMAT_RANGE');
      expect(action.format?.fillColor).toBe('#C6EFCE');

      const batch = normalizeTier1HeaderFormatActions(
        [{ type: 'FORMAT_MATCHING_ROWS' } as SheetAction],
        workbookContext,
        'header should be in bg red',
      );
      expect(batch[0]?.type).toBe('FORMAT_RANGE');
      expect(batch[0]?.format?.fillColor).toBe('#FF6B6B');
    });

    it('keeps conditional Pending highlight as FORMAT_MATCHING_ROWS', () => {
      const action: SheetAction = {
        type: 'FORMAT_MATCHING_ROWS',
        sheetName: 'Purchase Register',
        range: 'A1:L51',
        hasHeaders: true,
        filter: {
          column: 'Payment Status',
          operator: 'equals',
          value: 'Pending',
        },
        format: { fillColor: '#FFC7CE' },
      };
      const out = normalizeTier1ConditionalFormatActions(
        [action],
        workbookContext,
        'highlight rows where Payment Status is Pending',
      );
      expect(out[0]?.type).toBe('FORMAT_MATCHING_ROWS');
      expect(out[0]?.filter?.value).toBe('Pending');
    });
  });
});
