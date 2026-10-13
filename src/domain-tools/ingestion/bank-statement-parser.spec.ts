import * as fs from 'fs';
import * as path from 'path';
import {
  BankStatementParseError,
  BankStatementParseResult,
  parseBankStatement,
} from './bank-statement-parser';
import { verifyBankStatement } from './bank-statement-verifier';
import { RawTable, RawTableRow } from './raw-table.types';

/** A sheet-style table: one string per column, `row N` refs like the task pane produces. */
function grid(rows: string[][], fileName = 'statement.csv'): RawTable {
  return {
    source: 'csv',
    fileName,
    layout: 'grid',
    rows: rows.map((cells, index) => ({
      ref: `row ${index + 1}`,
      cells: cells.map((t) => ({ t })),
    })),
  };
}

function loadFederalFixture(): RawTable {
  const file = path.join(__dirname, 'fixtures', 'synthetic-federal-bank-statement.rawtable.json');
  return JSON.parse(fs.readFileSync(file, 'utf8')) as RawTable;
}

function errorCode(table: RawTable): string {
  try {
    parseBankStatement(table);
  } catch (error) {
    expect(error).toBeInstanceOf(BankStatementParseError);
    return (error as BankStatementParseError).code;
  }
  throw new Error('expected parseBankStatement to throw');
}

/**
 * Layout taken from a real Federal Bank savings-account PDF: positions, the
 * three-line header, page footers, the totals line and the trailer are the
 * bank's own. Every name, number and amount is synthetic.
 * Root/ATTACHMENT_EXTRACTION_PLAN.md §6.10.
 */
describe('parseBankStatement — Federal Bank PDF layout', () => {
  let table: RawTable;
  let parsed: BankStatementParseResult;

  beforeAll(() => {
    table = loadFederalFixture();
    parsed = parseBankStatement(table);
  });

  it('reads every transaction and nothing else', () => {
    expect(parsed.rows).toHaveLength(54);
    expect(parsed.skipped).toEqual([]);
    expect(parsed.amountLayout).toBe('debit_credit');
    expect(parsed.hasBalanceColumn).toBe(true);
    expect(parsed.rows[0].date).toBe('2025-04-01');
    expect(parsed.rows[53].date).toBe('2025-04-27');
  });

  it('reads the first row in full', () => {
    expect(parsed.rows[0]).toEqual({
      date: '2025-04-01',
      valueDate: '2025-04-01',
      description: 'NFT/SAMPLE PAYER LTD/TESTN40271339933/TEST BANK',
      refNo: 'S67133557',
      amount: 8495.74,
      type: 'credit',
      balance: 13495.74,
      flags: [],
      sourceRowRef: {
        documentType: 'bank_statement',
        documentId: 'synthetic-federal-bank-statement.pdf',
        rowOrLine: 'p1 l20',
      },
    });
  });

  it('takes the opening balance and totals the statement prints', () => {
    expect(parsed.openingBalance).toBe(5000);
    expect(parsed.statedTotalDebit).toBe(46432.01);
    expect(parsed.statedTotalCredit).toBe(321365.5);
  });

  it('masks the account number', () => {
    expect(parsed.accountNumberMasked).toBe('XXXXXXXXXX1234');
  });

  it('is not fooled by the account-details line that mentions a date and a balance', () => {
    // "Effective Available Balance ... Date of Issue" sits above the real header.
    expect(parsed.headerRef).toBe('p1 l17');
  });

  it('keeps a six-digit deposit in the deposits column', () => {
    // Right-aligned, so it starts to the left of the "Deposits" heading. A
    // left-edge rule files this row as a withdrawal.
    const header = table.rows.find((row) => row.ref === 'p1 l17') as RawTableRow;
    const deposits = header.cells.find((cell) => cell.t === 'Deposits');
    const big = table.rows.flatMap((row) => row.cells).find((cell) => cell.t === '250000.00');
    expect(big?.x0).toBeLessThan(deposits?.x0 as number);

    const row = parsed.rows[24];
    expect(row).toMatchObject({
      date: '2025-04-13',
      description: 'CASH:SELF',
      amount: 250000,
      type: 'credit',
      balance: 258687.44,
    });
  });

  it('rejoins a description that wraps mid-word without adding a space', () => {
    expect(parsed.rows[1].description).toBe('UPIOUT/257637335544/sample.payee-1@testupi/spl/0000');
  });

  it('rejoins a description that wraps between words with a space', () => {
    const atm = parsed.rows.find((row) => row.description.startsWith('TO ATM/'));
    expect(atm?.description).toBe('TO ATM/589418309507/ABC SAMPLE TOWN\\ABC');
  });

  it('keeps a value date that differs from the transaction date', () => {
    expect(parsed.rows[4]).toMatchObject({ date: '2025-04-03', valueDate: '2025-04-04' });
    expect(parsed.rows[19]).toMatchObject({ date: '2025-04-10', valueDate: '2025-04-11' });
  });

  it('does not read the DR/CR balance-sign column as the transaction direction', () => {
    // Every row says "Cr" there; 38 of the 54 are withdrawals.
    expect(parsed.rows.filter((row) => row.type === 'debit')).toHaveLength(38);
    expect(parsed.rows.filter((row) => row.type === 'credit')).toHaveLength(16);
  });

  it('never puts footer or trailer text into a description', () => {
    for (const row of parsed.rows) {
      expect(row.description).not.toMatch(/Federal Bank|Page \d|Ph:|DISCLAIMER|GRAND TOTAL|Abbreviations/);
    }
  });

  it('verifies: every row ties and the totals match the statement', () => {
    const { rows, verification } = verifyBankStatement(parsed);
    expect(verification).toEqual({
      state: 'verified',
      notes: [],
      order: 'oldest_first',
      rowsChecked: 54,
      rowsFailed: 0,
      rowsFlagged: 0,
      directionFixes: 0,
      openingBalance: 5000,
      openingBalanceSource: 'stated',
      closingBalance: 279933.49,
      totalDebit: 46432.01,
      totalCredit: 321365.5,
      totalsCheck: 'match',
    });
    expect(rows.every((row) => row.flags.length === 0)).toBe(true);
  });

  it('names the row when one balance is corrupted', () => {
    const broken: RawTable = JSON.parse(JSON.stringify(table));
    const target = broken.rows.find((row) => row.ref === 'p2 l1') as RawTableRow;
    const balance = target.cells[target.cells.length - 2];
    balance.t = (Number(balance.t) + 10).toFixed(2);

    const { rows, verification } = verifyBankStatement(parseBankStatement(broken));
    expect(verification.state).toBe('verified_with_exceptions');
    // The bad row fails, and so does the row after it, whose own balance is right.
    expect(verification.rowsFailed).toBe(2);
    const flagged = rows.filter((row) => row.flags.length > 0);
    expect(flagged.map((row) => row.sourceRowRef.rowOrLine)).toEqual(['p2 l1', 'p2 l3']);
    expect(flagged[0].flags[0]).toMatch(/^Balance does not tie: expected \d+\.\d{2}, statement shows \d+\.\d{2}$/);
  });

  it('reports a row that sits below the total line instead of importing or dropping it silently', () => {
    const extended: RawTable = JSON.parse(JSON.stringify(table));
    const totalIndex = extended.rows.findIndex((row) => row.cells.some((cell) => cell.t === 'GRAND TOTAL'));
    const template = extended.rows.find((row) => row.ref === 'p3 l1') as RawTableRow;
    extended.rows.splice(totalIndex + 1, 0, { ...template, ref: 'p3 l99', y: 600 });

    const result = parseBankStatement(extended);
    expect(result.rows).toHaveLength(54);
    expect(result.skipped).toEqual([expect.objectContaining({ ref: 'p3 l99', reason: 'after_total' })]);
    expect(verifyBankStatement(result).verification.state).toBe('verified_with_exceptions');
  });
});

describe('parseBankStatement — sheet layouts', () => {
  it('reads separate debit and credit columns below a block of account details', () => {
    const parsed = parseBankStatement(
      grid([
        ['Sample Bank Ltd'],
        ['Account No', '000111222333'],
        ['Statement Date', '30/04/2026', 'Available Balance', '1,05,000.00'],
        [],
        ['Txn Date', 'Value Date', 'Narration', 'Chq./Ref.No.', 'Withdrawal Amt.', 'Deposit Amt.', 'Closing Balance'],
        ['12/04/2026', '12/04/2026', 'NEFT IN SYNTH CORP', 'N0001', '', '15,000.00', '1,25,000.00'],
        ['13/04/2026', '13/04/2026', 'CHQ PAID RENT', '004521', '20,000.00', '', '1,05,000.00'],
      ]),
    );

    expect(parsed.headerRef).toBe('row 5');
    expect(parsed.accountNumberMasked).toBe('XXXXXXXX2333');
    expect(parsed.rows).toEqual([
      {
        date: '2026-04-12',
        valueDate: '2026-04-12',
        description: 'NEFT IN SYNTH CORP',
        refNo: 'N0001',
        amount: 15000,
        type: 'credit',
        balance: 125000,
        flags: [],
        sourceRowRef: { documentType: 'bank_statement', documentId: 'statement.csv', rowOrLine: 'row 6' },
      },
      {
        date: '2026-04-13',
        valueDate: '2026-04-13',
        description: 'CHQ PAID RENT',
        refNo: '004521',
        amount: 20000,
        type: 'debit',
        balance: 105000,
        flags: [],
        sourceRowRef: { documentType: 'bank_statement', documentId: 'statement.csv', rowOrLine: 'row 7' },
      },
    ]);
    // No opening balance is printed, so it is worked back from the first row.
    expect(verifyBankStatement(parsed).verification).toMatchObject({
      state: 'verified',
      openingBalance: 110000,
      openingBalanceSource: 'derived',
      rowsChecked: 1,
    });
  });

  it('reads one amount column with a Dr/Cr column', () => {
    const parsed = parseBankStatement(
      grid([
        ['Date', 'Description', 'Amount', 'Dr/Cr', 'Balance'],
        ['01-Apr-2026', 'SALARY', '50,000.00', 'CR', '60,000.00'],
        ['02-Apr-2026', 'RENT', '20,000.00', 'DR', '40,000.00'],
      ]),
    );
    expect(parsed.amountLayout).toBe('amount_drcr');
    expect(parsed.rows.map((row) => [row.type, row.amount, row.balance])).toEqual([
      ['credit', 50000, 60000],
      ['debit', 20000, 40000],
    ]);
  });

  it('reads one signed amount column', () => {
    const parsed = parseBankStatement(
      grid([
        ['Date', 'Description', 'Amount', 'Balance'],
        ['01-Apr-2026', 'SALARY', '50000.00', '60000.00'],
        ['02-Apr-2026', 'RENT', '-20000.00', '40000.00'],
        ['03-Apr-2026', 'CARD', '(500.00)', '39500.00'],
      ]),
    );
    expect(parsed.amountLayout).toBe('signed_amount');
    expect(parsed.rows.map((row) => [row.type, row.amount])).toEqual([
      ['credit', 50000],
      ['debit', 20000],
      ['debit', 500],
    ]);
    expect(verifyBankStatement(parsed).verification.state).toBe('verified');
  });

  it('reads amounts and balances that carry a Dr/Cr suffix', () => {
    const parsed = parseBankStatement(
      grid([
        ['Date', 'Particulars', 'Amount', 'Balance'],
        ['01-Apr-2026', 'OPENING TRANSFER', '1,000.00 Cr', '500.00 Dr'],
        ['02-Apr-2026', 'CHARGES', '200.00 Dr', '700.00 Dr'],
        ['03-Apr-2026', 'DEPOSIT', '1,000.00Cr', '300.00 Cr'],
      ]),
    );
    expect(parsed.rows.map((row) => [row.type, row.amount, row.balance])).toEqual([
      ['credit', 1000, -500],
      ['debit', 200, -700],
      ['credit', 1000, 300],
    ]);
    expect(verifyBankStatement(parsed).verification).toMatchObject({ state: 'verified', rowsChecked: 2 });
  });

  it('treats a "Type" column as Dr/Cr only when its values are Dr/Cr', () => {
    const parsed = parseBankStatement(
      grid([
        ['Date', 'Particulars', 'Type', 'Withdrawals', 'Deposits', 'Balance'],
        ['01-Apr-2026', 'TRANSFER IN', 'TFR', '', '100.00', '100.00'],
        ['02-Apr-2026', 'CASH OUT', 'CASH', '40.00', '', '60.00'],
      ]),
    );
    expect(parsed.rows.map((row) => [row.type, row.balance])).toEqual([
      ['credit', 100],
      ['debit', 60],
    ]);
  });

  it('merges continuation rows and uses the opening, closing and total lines', () => {
    const parsed = parseBankStatement(
      grid([
        ['Date', 'Narration', 'Debit', 'Credit', 'Balance'],
        ['', 'Opening Balance', '', '', '1,000.00'],
        ['01/04/2026', 'UPI PAYMENT TO', '250.00', '', '750.00'],
        ['', 'TOTAL ENERGIES FUEL STATION', '', '', ''],
        ['02/04/2026', 'INTEREST', '', '12.50', '762.50'],
        ['', 'Page Total', '250.00', '12.50', ''],
        ['', 'Closing Balance', '', '', '762.50'],
        ['', 'Total', '250.00', '12.50', ''],
      ]),
    );
    expect(parsed.rows.map((row) => row.description)).toEqual([
      'UPI PAYMENT TO TOTAL ENERGIES FUEL STATION',
      'INTEREST',
    ]);
    expect(parsed).toMatchObject({
      openingBalance: 1000,
      closingBalance: 762.5,
      statedTotalDebit: 250,
      statedTotalCredit: 12.5,
      skipped: [],
    });
    expect(verifyBankStatement(parsed).verification).toMatchObject({
      state: 'verified',
      rowsChecked: 2,
      openingBalanceSource: 'stated',
      totalsCheck: 'match',
    });
  });

  it('skips a header row repeated further down', () => {
    const header = ['Date', 'Narration', 'Debit', 'Credit', 'Balance'];
    const parsed = parseBankStatement(
      grid([
        header,
        ['01/04/2026', 'ONE', '10.00', '', '90.00'],
        header,
        ['02/04/2026', 'TWO', '10.00', '', '80.00'],
      ]),
    );
    expect(parsed.rows.map((row) => row.description)).toEqual(['ONE', 'TWO']);
    expect(parsed.skipped).toEqual([]);
  });

  it('reads day-first dates, and month-first ones when the column says so', () => {
    const table = (dates: string[]) =>
      grid([
        ['Date', 'Narration', 'Debit', 'Credit', 'Balance'],
        ...dates.map((date, index) => [date, `ROW ${index}`, '1.00', '', `${99 - index}.00`]),
      ]);
    expect(parseBankStatement(table(['05/04/2026', '06/04/2026'])).rows.map((row) => row.date)).toEqual([
      '2026-04-05',
      '2026-04-06',
    ]);
    expect(parseBankStatement(table(['04/05/2026', '04/25/2026'])).rows.map((row) => row.date)).toEqual([
      '2026-04-05',
      '2026-04-25',
    ]);
  });

  it('accepts ISO dates and plain numbers, as an Excel file is decoded', () => {
    const parsed = parseBankStatement(
      grid(
        [
          ['Transaction Date', 'Description', 'Debit', 'Credit', 'Balance'],
          ['2026-04-12', 'NEFT IN', '', '15000', '125000'],
          ['2026-04-13', 'RENT', '20000.5', '', '104999.5'],
        ],
        'statement.xlsx',
      ),
    );
    expect(parsed.rows.map((row) => [row.date, row.type, row.amount, row.balance])).toEqual([
      ['2026-04-12', 'credit', 15000, 125000],
      ['2026-04-13', 'debit', 20000.5, 104999.5],
    ]);
  });

  it('keeps and flags a row it cannot read, never dropping it', () => {
    const parsed = parseBankStatement(
      grid([
        ['Date', 'Narration', 'Debit', 'Credit', 'Balance'],
        ['01/04/2026', 'GOOD', '10.00', '', '90.00'],
        ['02/04/2026', 'NO AMOUNT', '', '', '90.00'],
        ['03/04/2026', 'BAD AMOUNT', 'ten', '', '80.00'],
        ['04/04/2026', 'BOTH SIDES', '5.00', '5.00', '80.00'],
      ]),
    );
    expect(parsed.rows.map((row) => row.flags)).toEqual([
      [],
      ['No amount on this row'],
      ['Debit not read: "ten"', 'No amount on this row'],
      ['Both a debit and a credit on this row'],
    ]);
    expect(verifyBankStatement(parsed).verification.state).toBe('verified_with_exceptions');
  });

  it('reports a line it does not recognise', () => {
    const parsed = parseBankStatement(
      grid([
        ['Date', 'Narration', 'Debit', 'Credit', 'Balance'],
        ['01/04/2026', 'ONE', '10.00', '', '90.00'],
        ['', '', 'stray text', '', ''],
        ['02/04/2026', 'TWO', '10.00', '', '80.00'],
      ]),
    );
    expect(parsed.rows).toHaveLength(2);
    expect(parsed.skipped).toEqual([{ ref: 'row 3', text: 'stray text', reason: 'unrecognised' }]);
  });
});

describe('parseBankStatement — refusals', () => {
  it('refuses an empty table', () => {
    expect(errorCode(grid([]))).toBe('empty');
  });

  it('refuses a file with no transaction table', () => {
    expect(
      errorCode(
        grid([
          ['Invoice No', 'Customer', 'Total'],
          ['INV-1', 'Acme', '100.00'],
        ]),
      ),
    ).toBe('no_header');
  });

  it('does not accept a loose mention of a date and a balance as the header', () => {
    expect(
      errorCode(
        grid([
          ['Effective Available Balance', '56506.25', 'Date of Issue', '01/10/2026'],
          ['Account Open Date', '01/01/2020', 'Account Status', 'ACTIVE'],
        ]),
      ),
    ).toBe('no_header');
  });

  it('refuses a table whose only money column is the balance', () => {
    expect(
      errorCode(
        grid([
          ['Date', 'Narration', 'Balance'],
          ['01/04/2026', 'ONE', '90.00'],
        ]),
      ),
    ).toBe('no_amount_columns');
  });

  it('refuses a header with no rows under it', () => {
    expect(errorCode(grid([['Date', 'Narration', 'Debit', 'Credit', 'Balance']]))).toBe('no_transactions');
  });
});
