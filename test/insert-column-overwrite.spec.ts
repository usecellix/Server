import { normalizeExecutorOutput } from '../src/agents/utils/normalize-executor-output.util';
import { ConversationEngineService } from '../src/excel-ai/services/conversation-engine.service';
import { SheetAnalysis } from '../src/excel-ai/services/sheet-analyzer.service';
import { virtualApply } from '../src/virtual/virtualApply';
import { ShadowWorkbook } from '../src/virtual/shadowWorkbook.types';
import { buildExecutorUserMessage, EXECUTOR_SYSTEM_PROMPT } from '../src/agents/prompts/executor.prompt';

describe('INSERT_COLUMN semantic + overwrite safety (spec 14)', () => {
  const analysis: SheetAnalysis = {
    rowCount: 52,
    columnCount: 12,
    headers: [
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
    ],
    isEmpty: false,
    columnLetters: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'],
    headerRowIndex: 0,
  };

  const service = new ConversationEngineService(
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
  );

  it('normalizes and sanitizes INSERT_COLUMN afterLastColumn without col', () => {
    const normalized = normalizeExecutorOutput(
      {
        subtaskId: 's1',
        actions: [
          {
            type: 'INSERT_COLUMN',
            sheetName: 'Purchase Register',
            columnName: 'Net of Tax',
            position: 'afterLastColumn',
            formula: '=J{row}-I{row}',
          },
        ],
        isDone: true,
      },
      { id: 's1', description: 'Add Net of Tax', targetSheet: 'Purchase Register', dependsOn: [], estimatedActions: 1 },
    );

    expect(normalized.actions).toHaveLength(1);
    expect(normalized.actions[0]).toMatchObject({
      type: 'INSERT_COLUMN',
      columnName: 'Net of Tax',
      position: 'afterLastColumn',
      formula: '=J{row}-I{row}',
    });

    const finalized = service.finalizeActions(normalized.actions, analysis);
    expect(finalized).toHaveLength(1);
    expect(finalized[0].type).toBe('INSERT_COLUMN');
    expect(finalized[0].columnName).toBe('Net of Tax');
    expect(finalized[0].col).toBeUndefined();
  });

  it('normalizes position.afterColumn object into afterColumn field', () => {
    const normalized = normalizeExecutorOutput(
      {
        subtaskId: 's1',
        actions: [
          {
            type: 'INSERT_COLUMN',
            sheetName: 'Purchase Register',
            columnName: 'Net of Tax',
            position: { afterColumn: 'Total Amount' },
            formula: '=J{row}-I{row}',
          },
        ],
        isDone: true,
      },
      { id: 's1', description: 'Add Net of Tax', targetSheet: 'Purchase Register', dependsOn: [], estimatedActions: 1 },
    );

    expect(normalized.actions[0]).toMatchObject({
      type: 'INSERT_COLUMN',
      columnName: 'Net of Tax',
      afterColumn: 'Total Amount',
      formula: '=J{row}-I{row}',
    });
  });

  it('virtualApply afterLastColumn writes into next empty column (M), leaving K untouched', () => {
    const wb: ShadowWorkbook = {
      activeSheetName: 'Purchase Register',
      sheets: new Map(),
      namedRanges: new Map(),
      tables: [],
      changedCells: new Set(),
    };
    const sheet = {
      name: 'Purchase Register',
      cells: new Map(),
      rowCount: 3,
      columnCount: 12,
      structure: 'data_table',
    };
    // Headers A-L
    const headers = analysis.headers;
    headers.forEach((h, i) => {
      const letter = String.fromCharCode(65 + i);
      sheet.cells.set(`${letter}1`, { value: h, formula: '', numberFormat: 'General' });
    });
    // Payment Status in K2/K3
    sheet.cells.set('K2', { value: 'Paid', formula: '', numberFormat: 'General' });
    sheet.cells.set('K3', { value: 'Pending', formula: '', numberFormat: 'General' });
    // Amounts for formula refs
    sheet.cells.set('I2', { value: 10, formula: '', numberFormat: 'General' });
    sheet.cells.set('J2', { value: 110, formula: '', numberFormat: 'General' });
    wb.sheets.set(sheet.name, sheet);

    const after = virtualApply(wb, [
      {
        type: 'INSERT_COLUMN',
        sheetName: 'Purchase Register',
        columnName: 'Net of Tax',
        position: 'afterLastColumn',
        formula: '=J{row}-I{row}',
      },
    ]);

    const result = after.sheets.get('Purchase Register')!;
    expect(result.cells.get('K2')?.value).toBe('Paid');
    expect(result.cells.get('K3')?.value).toBe('Pending');
    expect(result.cells.get('M1')?.value).toBe('Net of Tax');
    expect(result.cells.get('M2')?.formula).toBe('=J2-I2');
    expect(result.cells.get('M3')?.formula).toBe('=J3-I3');
  });

  it('executor prompt requires INSERT_COLUMN for add-column requests', () => {
    expect(EXECUTOR_SYSTEM_PROMPT).toContain('INSERT_COLUMN schema');
    expect(EXECUTOR_SYSTEM_PROMPT).toContain('afterLastColumn');
    expect(EXECUTOR_SYSTEM_PROMPT).toContain('ADD COLUMN (critical)');
    expect(EXECUTOR_SYSTEM_PROMPT).toMatch(/NEVER target an existing column with SET_CELL/);

    const message = buildExecutorUserMessage(
      {
        id: 's1',
        description: 'Add a column called Net of Tax that subtracts Tax Amount from Total Amount',
        targetSheet: 'Purchase Register',
        suggestedActionType: 'INSERT_COLUMN',
        dependsOn: [],
        estimatedActions: 1,
      },
      {
        activeSheetName: 'Purchase Register',
        sheets: [
          {
            name: 'Purchase Register',
            rowCount: 52,
            columnCount: 12,
            usedRange: 'A1:L52',
            structure: 'data_table',
            values: [analysis.headers],
            formulas: [[]],
            numberFormats: [[]],
          },
        ],
        namedRanges: [],
        tables: [],
        onDemandFetchEnabled: true,
      },
      [],
    );
    expect(message).toContain('Suggested action type: INSERT_COLUMN');
  });
});
