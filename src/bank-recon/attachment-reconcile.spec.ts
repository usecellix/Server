import { BankReconService } from './bank-recon.service';
import { BankStatementService } from '../bank-statement/bank-statement.service';

/**
 * Reconcile the employee's ledger sheet against an ATTACHED statement — the decoded file goes
 * straight to the reconciliation, never into a sheet first. The fixtures are the exact tables
 * the task pane's decoder produces for the demo files (CSV with Dr/Cr + BOM, semicolon CSV in
 * cp1252, Excel, digital PDF), so this covers the whole read path after decoding.
 */
// eslint-disable-next-line @typescript-eslint/no-require-imports
const fx = require('./__fixtures__/demo-sep2026.json') as Record<string, unknown[][]>;
const att = (name: string) =>
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  require(`./__fixtures__/attachments/${name}.json`) as Record<string, unknown> & { fileName: string };

const reconcile = (rawTable: Record<string, unknown>, extra: Record<string, unknown> = {}) =>
  new BankReconService().reconcile({
    books_ledger: { sheet_name: 'Bank Ledger', data: fx.sheet1 },
    bank_statement: { sheet_name: String(rawTable.fileName), raw_table: rawTable },
    prior_brs: { sheet_name: 'Prior BRS (Aug-26)', data: fx.sheet3 },
    casual: true,
    layout: 'categorized',
    ...extra,
  } as never) as any;

const FORMATS = [
  'bank_statement_sep2026_A_drcr_utf8bom.csv',
  'bank_statement_sep2026_B_semicolon_cp1252.csv',
  'bank_statement_sep2026.xlsx',
  'bank_statement_sep2026_digital.pdf',
];

describe('reconcile against an attached statement', () => {
  it.each(FORMATS)('%s → the same reconciliation as the sheet-based run', (name) => {
    const r = reconcile(att(name));
    expect(r.bank_source).toMatchObject({ kind: 'attachment', transactions: 65, verification: 'verified' });
    expect(r.summary).toMatchObject({ matched: 43, matched_group: 2, books_only: 10, bank_only: 20 });
    expect(r.brs.tiedOut).toBe(true);
    expect(r.brs.bankBalance).toBe(2478981.05);
    expect(r.validation.errors).toEqual([]);
  });

  it('bank-side rows point into the file (no workbook cell), books-side rows still point at the ledger sheet', () => {
    const r = reconcile(att('bank_statement_sep2026_digital.pdf'));
    const bank = r.rows.filter((x: any) => x.side === 'BANK');
    const books = r.rows.filter((x: any) => x.side === 'BOOKS');
    expect(bank.length).toBe(20);
    expect(bank.every((x: any) => x.external === true && /^p\d+ l\d+$/.test(x.source_ref))).toBe(true);
    expect(books.every((x: any) => !x.external)).toBe(true);

    // The sheet's "Source Row" column names the file and the place in it, not a sheet!row.
    const table = r.actions.find((a: any) => a.type === 'WRITE_TABLE') as { rows: unknown[][] };
    const sourceCells = table.rows.map((row) => String(row[8] ?? '')).filter((c) => c.includes('bank_statement_sep2026_digital.pdf'));
    expect(sourceCells.length).toBe(20);
    expect(sourceCells[0]).toMatch(/digital\.pdf · p\d+ l\d+/);
  });

  it('refuses a statement that was read incompletely instead of reporting false mismatches (hybrid PDF with an image-only page)', () => {
    expect(() => reconcile(att('bank_statement_sep2026_hybrid.pdf'))).toThrow(/could not be read completely/);
  });

  it('refuses a file that is not a statement, with a message the user can act on', () => {
    const notAStatement = {
      source: 'csv',
      layout: 'grid',
      fileName: 'notes.csv',
      rows: [{ ref: 'row 1', cells: [{ t: 'hello' }, { t: 'world' }] }],
    };
    expect(() => reconcile(notAStatement)).toThrow();
  });

  it('still needs a statement: neither a sheet nor an attached file is a 400', () => {
    expect(() =>
      new BankReconService().reconcile({
        books_ledger: { sheet_name: 'Bank Ledger', data: fx.sheet1 },
        bank_statement: { sheet_name: 'x' },
      } as never),
    ).toThrow(/required/);
  });
});

/**
 * The "Add as sheet" choice: the statement importer writes the file into a sheet first, then the reconciliation
 * reads THAT sheet (rows clickable). It must give the very same reconciliation as reading the file directly,
 * for every format — including PDF.
 */
describe("'Add as sheet' path — import the file into a sheet, then reconcile against the sheet", () => {
  const ALL = [...FORMATS, 'bank_statement_sep2026_hybrid.pdf'];

  const importedSheetGrid = (name: string) => {
    const imp = new BankStatementService().importStatement({ rawTable: att(name) as never, existingSheetNames: [] });
    const write = imp.actions.find((a: any) => a.type === 'WRITE_TABLE') as { headers: string[]; rows: unknown[][] };
    return { grid: [write.headers, ...write.rows] as unknown[][], imp };
  };

  it.each(FORMATS)('%s → same reconciliation as reading the file directly', (name) => {
    const direct = reconcile(att(name));
    const { grid, imp } = importedSheetGrid(name);
    expect(imp.statement.transactionCount).toBe(65);

    const viaSheet = new BankReconService().reconcile({
      books_ledger: { sheet_name: 'Bank Ledger', data: fx.sheet1 },
      bank_statement: { sheet_name: imp.outputSheetName, data: grid },
      prior_brs: { sheet_name: 'Prior BRS (Aug-26)', data: fx.sheet3 },
      casual: true,
      layout: 'categorized',
    } as never) as any;

    expect(viaSheet.summary).toMatchObject({ matched: direct.summary.matched, matched_group: 2, books_only: 10, bank_only: 20 });
    expect(viaSheet.brs.tiedOut).toBe(true);
    expect(viaSheet.brs.bankBalance).toBe(direct.brs.bankBalance);
    // with a sheet, every bank-side row is a clickable sheet row (row 2 onward — row 1 is the header)
    const bank = viaSheet.rows.filter((r: any) => r.side === 'BANK');
    expect(bank).toHaveLength(20);
    expect(bank.every((r: any) => !r.external && r.sheet_name === imp.outputSheetName && r.row >= 2 && r.row <= 66)).toBe(true);
  });

  it('an incompletely-read PDF that was added as a sheet is refused just like the file itself (strict), but a plain sheet is not', () => {
    const { grid, imp } = importedSheetGrid('bank_statement_sep2026_hybrid.pdf');
    const run = (strict?: boolean) =>
      new BankReconService().reconcile({
        books_ledger: { sheet_name: 'Bank Ledger', data: fx.sheet1 },
        bank_statement: { sheet_name: imp.outputSheetName, data: grid, ...(strict ? { strict: true } : {}) },
        prior_brs: { sheet_name: 'Prior BRS (Aug-26)', data: fx.sheet3 },
        casual: true,
      } as never) as any;

    expect(() => run(true)).toThrow(/could not be read completely/);
    // a sheet the user typed or pasted themselves is theirs to review: warnings, not a refusal
    const plain = run(false);
    expect(plain.validation.errors.length).toBeGreaterThan(0);
  });

  it('a PDF read in the import is flagged, not silently trusted, when part of it could not be read (hybrid)', () => {
    const { imp } = importedSheetGrid('bank_statement_sep2026_hybrid.pdf');
    expect(imp.statement.transactionCount).toBeLessThan(65);
    expect(imp.verification.state).not.toBe('verified');
    expect(ALL).toContain('bank_statement_sep2026_hybrid.pdf');
  });
});
