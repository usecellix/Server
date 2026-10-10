import { NormalizedBankStatementRow } from '../domain-tools/ingestion/bank-statement-parser';
import {
  BANK_STATEMENT_HEADERS,
  bankStatementSheetRow,
  mapBankStatementToActions,
  pickBankStatementSheetName,
} from './bank-statement-to-actions.mapper';

function row(overrides: Partial<NormalizedBankStatementRow> = {}): NormalizedBankStatementRow {
  return {
    date: '2026-04-11',
    valueDate: '2026-04-12',
    description: 'NEFT IN SYNTH CORP',
    refNo: 'S12345678',
    amount: 1627,
    type: 'credit',
    balance: 2544.04,
    flags: [],
    sourceRowRef: { documentType: 'bank_statement', documentId: 'statement.pdf', rowOrLine: 'p1 l20' },
    ...overrides,
  };
}

describe('bankStatementSheetRow', () => {
  it('lays a credit out in the fixed nine columns', () => {
    expect(bankStatementSheetRow(row())).toEqual([
      46123, // 11-Apr-2026 as an Excel date serial
      46124,
      'NEFT IN SYNTH CORP',
      'S12345678',
      '',
      1627,
      2544.04,
      'p1 l20',
      '',
    ]);
  });

  it('puts a debit in the Debit column', () => {
    const cells = bankStatementSheetRow(row({ type: 'debit', amount: 180 }));
    expect([cells[4], cells[5]]).toEqual([180, '']);
  });

  it('leaves blanks, never zeros or text, where the statement had nothing', () => {
    const cells = bankStatementSheetRow(
      row({ date: '', valueDate: undefined, refNo: undefined, balance: undefined, flags: ['Date not read: "??"'] }),
    );
    expect(cells).toEqual(['', '', 'NEFT IN SYNTH CORP', '', '', 1627, '', 'p1 l20', 'Date not read: "??"']);
  });

  it('keeps a negative balance negative', () => {
    expect(bankStatementSheetRow(row({ balance: -500 }))[6]).toBe(-500);
  });

  it('joins several flags', () => {
    expect(bankStatementSheetRow(row({ flags: ['No amount on this row', 'No balance on this row'] }))[8]).toBe(
      'No amount on this row; No balance on this row',
    );
  });

  it.each(['=HYPERLINK("http://x","y")', '+91 REFUND', '-REVERSAL', '@handle', '\tTAB'])(
    'writes %p as literal text, not a formula',
    (text) => {
      const cells = bankStatementSheetRow(row({ description: text, refNo: text }));
      expect(cells[2]).toBe(`'${text}`);
      expect(cells[3]).toBe(`'${text}`);
    },
  );
});

describe('mapBankStatementToActions', () => {
  const rows = [row(), row({ type: 'debit', amount: 180, refNo: '004521' })];
  const actions = mapBankStatementToActions({ sheetName: 'Bank Statement', rows, relativeTo: 'Sheet1' });
  const types = actions.map((action) => action.type);

  it('creates the sheet first and writes one table to it', () => {
    expect(actions[0]).toEqual({
      type: 'CREATE_SHEET',
      sheetName: 'Bank Statement',
      name: 'Bank Statement',
      relativeTo: 'Sheet1',
      position: 'after',
    });
    const writes = actions.filter((action) => action.type === 'WRITE_TABLE');
    expect(writes).toHaveLength(1);
    expect(writes[0].headers).toEqual([...BANK_STATEMENT_HEADERS]);
    expect(writes[0].rows).toHaveLength(2);
  });

  it('targets only the new sheet', () => {
    expect(new Set(actions.map((action) => action.sheetName))).toEqual(new Set(['Bank Statement']));
  });

  it('sets the text columns to Text before the write, so "004521" keeps its zeros', () => {
    const writeIndex = types.indexOf('WRITE_TABLE');
    const textFormats = actions
      .map((action, index) => ({ action, index }))
      .filter(({ action }) => action.type === 'FORMAT_RANGE' && action.format?.numberFormat === '@');
    expect(textFormats.map(({ action }) => action.range)).toEqual(['C2:D3', 'H2:I3']);
    expect(textFormats.every(({ index }) => index < writeIndex)).toBe(true);
  });

  it('formats dates and amounts after the write, over exactly the data rows', () => {
    const writeIndex = types.indexOf('WRITE_TABLE');
    const after = actions.slice(writeIndex + 1).filter((action) => action.type === 'FORMAT_RANGE');
    expect(after.map((action) => [action.range, action.format?.numberFormat])).toEqual([
      ['A1:I1', undefined],
      ['A2:B3', 'dd-mmm-yyyy'],
      ['E2:G3', '#,##0.00'],
    ]);
  });

  it('sets a width for each of the nine columns and freezes the header row', () => {
    const widths = actions.filter((action) => action.type === 'SET_COLUMN_WIDTH');
    expect(widths.map((action) => action.col)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8]);
    expect(actions[actions.length - 1]).toEqual({ type: 'FREEZE_PANES', sheetName: 'Bank Statement', freezeRows: 1 });
  });

  it('uses only action types that create or format, never one that edits existing cells', () => {
    expect(new Set(types)).toEqual(
      new Set(['CREATE_SHEET', 'FORMAT_RANGE', 'WRITE_TABLE', 'SET_COLUMN_WIDTH', 'FREEZE_PANES']),
    );
  });

  it('omits the position when no sheet is given to place it after', () => {
    expect(mapBankStatementToActions({ sheetName: 'Bank Statement', rows })[0]).toEqual({
      type: 'CREATE_SHEET',
      sheetName: 'Bank Statement',
      name: 'Bank Statement',
    });
  });
});

describe('pickBankStatementSheetName', () => {
  it('uses the plain name in a workbook that does not have it', () => {
    expect(pickBankStatementSheetName()).toBe('Bank Statement');
    expect(pickBankStatementSheetName(['Sheet1', 'Bank Recon'])).toBe('Bank Statement');
  });

  it('never returns a name the workbook already has', () => {
    expect(pickBankStatementSheetName(['Bank Statement'])).toBe('Bank Statement 2');
    expect(pickBankStatementSheetName(['Bank Statement', 'Bank Statement 2', 'Bank Statement 3'])).toBe(
      'Bank Statement 4',
    );
  });

  it('compares names the way Excel does, ignoring case and outer spaces', () => {
    expect(pickBankStatementSheetName(['bank statement ', 'BANK STATEMENT 2'])).toBe('Bank Statement 3');
  });
});
