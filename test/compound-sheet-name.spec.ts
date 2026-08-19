import {
  buildDeterministicSubtaskActions,
  extractQuotedSheetName,
  extractSheetNameFromSubtaskDescription,
  pruneSpuriousAddSheetActions,
  suggestNewSheetName,
} from '../src/agents/utils/compound-action.util';
import { WorkbookContext } from '../src/agents/types/agent.types';

const emptyContext = (): WorkbookContext => ({
  activeSheetName: 'Purchase Register',
  sheets: [
    {
      name: 'Purchase Register',
      rowCount: 51,
      columnCount: 12,
      usedRange: 'A1:L51',
      structure: 'data_table',
      headerRowIndex: 0,
      values: [[]],
      formulas: [[]],
      numberFormats: [[]],
    },
  ],
  namedRanges: [],
  tables: [],
  onDemandFetchEnabled: true,
});

describe('compound sheet name extraction (paid purchases / Sheet2 bug)', () => {
  it('extracts single-quoted sheet names from planner descriptions', () => {
    expect(
      extractQuotedSheetName("Create sheet 'paid paid purchases' if it doesn't exist"),
    ).toBe('paid paid purchases');
  });

  it('does not fall back to Sheet2 when the description names the sheet', () => {
    const ctx = emptyContext();
    const name = suggestNewSheetName(
      "Create sheet 'paid paid purchases' if it doesn't exist",
      ctx,
    );
    expect(name.toLowerCase()).toBe('paid paid purchases');
    expect(name.toLowerCase()).not.toBe('sheet2');
  });

  it('buildDeterministicSubtaskActions emits the named sheet, not Sheet2', () => {
    const result = buildDeterministicSubtaskActions(
      {
        id: 's1',
        description: "Create sheet 'paid paid purchases' if it doesn't exist",
        targetSheet: 'Purchase Register',
        dependsOn: [],
        estimatedActions: 1,
      },
      emptyContext(),
    );
    expect(result?.actions).toEqual([
      { type: 'ADD_SHEET', name: 'paid paid purchases' },
    ]);
  });

  it('prunes phantom Sheet2 when a real ADD_SHEET + COPY share the batch', () => {
    const pruned = pruneSpuriousAddSheetActions([
      { type: 'ADD_SHEET', name: 'Sheet2' },
      { type: 'ADD_SHEET', name: 'paid paid purchases' },
      {
        type: 'COPY_FILTERED_RANGE',
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:L51',
        destSheet: 'paid paid purchases',
        destStartCell: 'A1',
        hasHeaders: true,
        mode: 'copy',
        filter: { column: 'Payment Status', operator: 'equals', value: 'Paid' },
      },
    ]);
    expect(pruned.map((a) => a.type + ':' + (a.name ?? a.destSheet))).toEqual([
      'ADD_SHEET:paid paid purchases',
      'COPY_FILTERED_RANGE:paid paid purchases',
    ]);
  });

  it('extractSheetNameFromSubtaskDescription handles Create sheet X if…', () => {
    const name = extractSheetNameFromSubtaskDescription(
      'Create sheet PaidPurchases if it does not exist',
      emptyContext(),
    );
    expect(name).toBe('PaidPurchases');
  });

  it('skips COPY of header-only month sheets deterministically', () => {
    const ctx = emptyContext();
    ctx.sheets.push({
      name: 'January 2026',
      rowCount: 1,
      columnCount: 10,
      usedRange: 'A1:J1',
      structure: 'data_table',
      headerRowIndex: 0,
      values: [['Unit No', 'Guest']],
      formulas: [[]],
      numberFormats: [[]],
    });
    const result = buildDeterministicSubtaskActions(
      {
        id: 's12',
        description:
          "Copy all data rows (excluding header) from 'January 2026' to 'Main' starting at A21, appending below existing rows",
        targetSheet: 'Main',
        dependsOn: [],
        estimatedActions: 1,
        suggestedActionType: 'COPY_FILTERED_RANGE',
      },
      ctx,
    );
    expect(result).toEqual({
      subtaskId: 's12',
      actions: [],
      isDone: true,
    });
  });
});
