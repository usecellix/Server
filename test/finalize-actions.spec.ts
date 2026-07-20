import { ConversationEngineService } from '../src/excel-ai/services/conversation-engine.service';
import { SheetAnalysis } from '../src/excel-ai/services/sheet-analyzer.service';

describe('ConversationEngineService finalizeActions', () => {
  const service = new ConversationEngineService(
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
  );

  const analysis: SheetAnalysis = {
    rowCount: 339,
    columnCount: 13,
    headers: ['Date', 'Particulars', 'CGST'],
    isEmpty: false,
    columnLetters: ['A', 'B', 'C'],
    headerRowIndex: 0,
  };

  it('preserves SORT_RANGE actions through sanitization', () => {
    const result = service.finalizeActions(
      [
        {
          type: 'SORT_RANGE',
          sheetName: 'Purchases',
          range: 'A1:M339',
          key: 0,
          ascending: true,
          hasHeaders: true,
          columnName: 'Date',
        },
      ],
      analysis,
    );

    expect(result).toHaveLength(1);
    expect(result[0].type).toBe('SORT_RANGE');
    expect(result[0].sheetName).toBe('Purchases');
  });

  it('preserves COPY_FILTERED_RANGE through sanitization', () => {
    const result = service.finalizeActions(
      [
        {
          type: 'COPY_FILTERED_RANGE',
          sourceSheet: 'Purchase Register',
          sourceRange: 'A1:L51',
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
      ],
      analysis,
    );

    expect(result).toHaveLength(1);
    expect(result[0].type).toBe('COPY_FILTERED_RANGE');
    expect(result[0].destSheet).toBe('Pending Payments');
  });

  it('drops incomplete COPY_FILTERED_RANGE actions', () => {
    const result = service.finalizeActions(
      [
        {
          type: 'COPY_FILTERED_RANGE',
          sourceSheet: 'Purchase Register',
          // missing destSheet / destStartCell / mode
          sourceRange: 'A1:L51',
        },
      ],
      analysis,
    );

    expect(result).toHaveLength(0);
  });
});
