import * as fs from 'fs';
import * as path from 'path';
import { BadRequestException, UnprocessableEntityException } from '@nestjs/common';
import { RawTable } from '../domain-tools/ingestion/raw-table.types';
import { assertRawTable, BankStatementService } from './bank-statement.service';

const fixture = (): RawTable =>
  JSON.parse(
    fs.readFileSync(
      path.join(
        __dirname,
        '..',
        'domain-tools',
        'ingestion',
        'fixtures',
        'synthetic-federal-bank-statement.rawtable.json',
      ),
      'utf8',
    ),
  ) as RawTable;

function grid(rows: string[][]): RawTable {
  return {
    source: 'csv',
    fileName: 'statement.csv',
    layout: 'grid',
    rows: rows.map((cells, index) => ({ ref: `row ${index + 1}`, cells: cells.map((t) => ({ t })) })),
  };
}

describe('BankStatementService.importStatement', () => {
  const service = new BankStatementService();

  it('turns the Federal Bank fixture into a verified import', () => {
    const result = service.importStatement({ rawTable: fixture(), existingSheetNames: ['Sheet1'] });

    expect(result.outputSheetName).toBe('Bank Statement');
    expect(result.statement).toEqual({
      fileName: 'synthetic-federal-bank-statement.pdf',
      source: 'pdf',
      accountNumberMasked: 'XXXXXXXXXX1234',
      transactionCount: 54,
      periodStart: '2025-04-01',
      periodEnd: '2025-04-27',
    });
    expect(result.verification.state).toBe('verified');
    expect(result.exceptions).toEqual([]);
    expect(result.skipped).toEqual([]);
    expect(result.skippedCount).toBe(0);

    const write = result.actions.find((action) => action.type === 'WRITE_TABLE');
    expect(write?.rows).toHaveLength(54);
    // First transaction: 01-Apr-2025 credit of 8495.74, balance 13495.74.
    expect(write?.rows?.[0]).toEqual([
      45748,
      45748,
      'NFT/SAMPLE PAYER LTD/TESTN40271339933/TEST BANK',
      'S67133557',
      '',
      8495.74,
      13495.74,
      'p1 l20',
      '',
    ]);
  });

  it('never returns the full account number', () => {
    const result = service.importStatement({ rawTable: fixture() });
    expect(JSON.stringify(result)).not.toContain('99990100001234');
  });

  it('gives the new sheet a free name and places it after the active sheet', () => {
    const result = service.importStatement({
      rawTable: fixture(),
      existingSheetNames: ['Bank Statement', 'Ledger'],
      activeSheetName: 'Ledger',
    });
    expect(result.outputSheetName).toBe('Bank Statement 2');
    expect(result.actions[0]).toMatchObject({ type: 'CREATE_SHEET', name: 'Bank Statement 2', relativeTo: 'Ledger' });
    expect(result.actions.every((action) => action.sheetName === 'Bank Statement 2')).toBe(true);
  });

  it('lists flagged rows by the sheet row they land on', () => {
    const result = service.importStatement({
      rawTable: grid([
        ['Date', 'Narration', 'Debit', 'Credit', 'Balance'],
        ['01/04/2026', 'GOOD', '10.00', '', '90.00'],
        ['02/04/2026', 'WRONG BALANCE', '10.00', '', '85.00'],
      ]),
    });
    expect(result.verification.state).toBe('verified_with_exceptions');
    expect(result.exceptions).toEqual([
      {
        sheetRow: 3,
        source: 'row 3',
        reason: 'Balance does not tie: expected 80.00, statement shows 85.00',
      },
    ]);
    const write = result.actions.find((action) => action.type === 'WRITE_TABLE');
    expect(write?.rows?.[1]?.[8]).toBe('Balance does not tie: expected 80.00, statement shows 85.00');
  });

  it('refuses a file that is not a statement with a reason the user can act on', () => {
    expect.assertions(3);
    try {
      service.importStatement({ rawTable: grid([['Invoice No', 'Customer', 'Total'], ['INV-1', 'Acme', '100.00']]) });
    } catch (error) {
      expect(error).toBeInstanceOf(UnprocessableEntityException);
      const body = (error as UnprocessableEntityException).getResponse() as { code: string; message: string };
      expect(body.code).toBe('no_header');
      expect(body.message).toMatch(/Could not find the transaction table/);
    }
  });
});

describe('assertRawTable', () => {
  const valid = (): RawTable => grid([['Date', 'Narration']]);

  it('accepts a well-formed table', () => {
    expect(() => assertRawTable(valid())).not.toThrow();
    expect(() => assertRawTable(fixture())).not.toThrow();
  });

  it.each<[string, (table: Record<string, unknown>) => void]>([
    ['an unknown source', (t) => (t.source = 'docx')],
    ['an unknown layout', (t) => (t.layout = 'freeform')],
    ['a missing file name', (t) => delete t.fileName],
    ['rows that are not an array', (t) => (t.rows = 'nope')],
    ['a row with no ref', (t) => ((t.rows as Record<string, unknown>[])[0].ref = undefined)],
    ['a row whose cells are not an array', (t) => ((t.rows as Record<string, unknown>[])[0].cells = {})],
    ['a cell whose text is not a string', (t) => ((t.rows as { cells: unknown[] }[])[0].cells[0] = { t: 5 })],
    ['a cell with a non-numeric position', (t) => ((t.rows as { cells: unknown[] }[])[0].cells[0] = { t: 'x', x0: 'left' })],
    ['a row with a non-numeric height', (t) => ((t.rows as Record<string, unknown>[])[0].y = Number.NaN)],
  ])('rejects %s', (_label, corrupt) => {
    const table = valid() as unknown as Record<string, unknown>;
    corrupt(table);
    expect(() => assertRawTable(table)).toThrow(BadRequestException);
  });

  it.each([null, undefined, 'text', 42])('rejects %p', (value) => {
    expect(() => assertRawTable(value)).toThrow(BadRequestException);
  });
});
