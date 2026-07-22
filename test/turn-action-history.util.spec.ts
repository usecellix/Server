import {
  collectRecentTurnActionRecords,
  extractTurnActionRecords,
  formatTurnActionRecordsForExecutor,
  referencesPriorChartOrTable,
  resolvePriorSourceRange,
} from '../src/excel-ai/utils/turn-action-history.util';

describe('turn-action-history.util', () => {
  it('detects prior-context follow-up phrasing', () => {
    expect(referencesPriorChartOrTable('Also create a bar chart along with the current')).toBe(
      true,
    );
    expect(referencesPriorChartOrTable('update that chart to a pie')).toBe(true);
    expect(referencesPriorChartOrTable('create a chart of purchases')).toBe(false);
  });

  it('prefers latest CREATE_CHART sourceRange for follow-ups', () => {
    const records = extractTurnActionRecords([
      {
        type: 'AGGREGATE_TABLE',
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:L200',
        destSheet: 'Dashboard',
        destStartCell: 'A4',
        groupByColumn: 'Supplier',
        aggregations: [{ column: 'Amount', fn: 'sum', outputLabel: 'Spend' }],
      },
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Dashboard',
        sourceRange: 'A1:B10',
        chartId: 'Chart_spend',
      },
    ]);
    expect(resolvePriorSourceRange(records)).toEqual({
      sourceRange: 'A1:B10',
      sourceSheetName: 'Dashboard',
      chartId: 'Chart_spend',
    });
    expect(formatTurnActionRecordsForExecutor(records)).toContain('Chart_spend');
  });

  it('falls back to extracting from metadata.actions when turnActionRecords missing', () => {
    const collected = collectRecentTurnActionRecords([
      {
        role: 'assistant',
        metadata: {
          actions: [
            {
              type: 'CREATE_CHART',
              sheetName: 'Dashboard',
              sourceRange: 'A1:B8',
              chartId: 'Chart_x',
            },
          ],
        },
      },
    ]);
    expect(collected).toHaveLength(1);
    expect(collected[0].sourceRange).toBe('A1:B8');
  });
});
