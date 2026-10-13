import { BankStatementParseResult, NormalizedBankStatementRow } from './bank-statement-parser';
import { verifyBankStatement } from './bank-statement-verifier';

let line = 0;
function txn(
  type: 'credit' | 'debit',
  amount: number,
  balance: number | undefined,
  flags: string[] = [],
): NormalizedBankStatementRow {
  line += 1;
  return {
    date: '2026-04-01',
    description: `ROW ${line}`,
    amount,
    type,
    ...(balance !== undefined ? { balance } : {}),
    flags,
    sourceRowRef: { documentType: 'bank_statement', documentId: 'test', rowOrLine: `row ${line}` },
  };
}

function statement(
  rows: NormalizedBankStatementRow[],
  extra: Partial<BankStatementParseResult> = {},
): BankStatementParseResult {
  return {
    rows,
    amountLayout: 'debit_credit',
    hasBalanceColumn: rows.some((row) => row.balance !== undefined),
    skipped: [],
    headerRef: 'row 1',
    ...extra,
  };
}

beforeEach(() => {
  line = 0;
});

describe('verifyBankStatement', () => {
  it('verifies a statement whose every balance follows from the row before', () => {
    const { rows, verification } = verifyBankStatement(
      statement([txn('credit', 100, 1100), txn('debit', 250.5, 849.5), txn('credit', 0.5, 850)], {
        openingBalance: 1000,
      }),
    );
    expect(verification).toMatchObject({
      state: 'verified',
      notes: [],
      rowsChecked: 3,
      rowsFailed: 0,
      openingBalance: 1000,
      openingBalanceSource: 'stated',
      closingBalance: 850,
      totalDebit: 250.5,
      totalCredit: 100.5,
      totalsCheck: 'not_stated',
    });
    expect(rows.every((row) => row.flags.length === 0)).toBe(true);
  });

  it('does not lose paise to floating point over a long run', () => {
    // 0.1 + 0.2 style sums: 1000 rows of 0.10 must land on exactly 100.00.
    let balance = 0;
    const rows = Array.from({ length: 1000 }, () => {
      balance = Math.round((balance + 0.1) * 100) / 100;
      return txn('credit', 0.1, balance);
    });
    const { verification } = verifyBankStatement(statement(rows, { openingBalance: 0, statedTotalCredit: 100 }));
    expect(verification).toMatchObject({ state: 'verified', rowsFailed: 0, totalCredit: 100, totalsCheck: 'match' });
  });

  it('flags only the row that does not tie, then carries on from the printed balance', () => {
    const { rows, verification } = verifyBankStatement(
      statement([txn('credit', 100, 1100), txn('debit', 50, 1000), txn('debit', 100, 900)], {
        openingBalance: 1000,
      }),
    );
    expect(verification).toMatchObject({ state: 'verified_with_exceptions', rowsFailed: 1, rowsFlagged: 1 });
    expect(rows.map((row) => row.flags)).toEqual([
      [],
      ['Balance does not tie: expected 1050.00, statement shows 1000.00'],
      [],
    ]);
  });

  it('corrects a row whose amount sat under the wrong heading', () => {
    // The balance fell by 40, so this was a withdrawal whatever column it was read from.
    const { rows, verification } = verifyBankStatement(
      statement([txn('credit', 100, 1100), txn('credit', 40, 1060)], { openingBalance: 1000 }),
    );
    expect(rows[1].type).toBe('debit');
    expect(verification).toMatchObject({
      state: 'verified',
      directionFixes: 1,
      rowsFailed: 0,
      totalDebit: 40,
      totalCredit: 100,
    });
  });

  it('reads a statement that lists the newest transaction first', () => {
    const { rows, verification } = verifyBankStatement(
      statement([txn('debit', 100, 900), txn('debit', 50, 1000), txn('credit', 50, 1050)]),
    );
    expect(verification).toMatchObject({
      state: 'verified',
      order: 'newest_first',
      rowsChecked: 2,
      openingBalance: 1000,
      openingBalanceSource: 'derived',
      closingBalance: 900,
    });
    // Rows stay in the order the statement gave them.
    expect(rows.map((row) => row.description)).toEqual(['ROW 1', 'ROW 2', 'ROW 3']);
  });

  it('cannot check the first row when no opening balance is printed', () => {
    const { verification } = verifyBankStatement(statement([txn('credit', 100, 1100), txn('debit', 100, 1000)]));
    expect(verification).toMatchObject({ state: 'verified', rowsChecked: 1, openingBalanceSource: 'derived' });
  });

  it('is unverified when there is no balance column', () => {
    const { verification } = verifyBankStatement(statement([txn('credit', 100, undefined), txn('debit', 40, undefined)]));
    expect(verification).toMatchObject({
      state: 'unverified',
      rowsChecked: 0,
      openingBalanceSource: 'none',
      totalDebit: 40,
      totalCredit: 100,
    });
    expect(verification.notes).toEqual(['The statement has no balance column, so the rows could not be checked.']);
    expect(verification.openingBalance).toBeUndefined();
    expect(verification.closingBalance).toBeUndefined();
  });

  it('says so when a statement without balances still agrees with its own totals', () => {
    const { verification } = verifyBankStatement(
      statement([txn('credit', 100, undefined), txn('debit', 40, undefined)], {
        statedTotalDebit: 40,
        statedTotalCredit: 100,
      }),
    );
    expect(verification.state).toBe('unverified');
    expect(verification.notes[0]).toMatch(/totals do match the statement/);
  });

  it('reports totals that disagree with the statement', () => {
    const { verification } = verifyBankStatement(
      statement([txn('credit', 100, 1100), txn('debit', 100, 1000)], {
        openingBalance: 1000,
        statedTotalDebit: 100,
        statedTotalCredit: 150,
      }),
    );
    expect(verification).toMatchObject({ state: 'verified_with_exceptions', totalsCheck: 'mismatch' });
    expect(verification.notes[0]).toMatch(/do not equal the imported rows/);
  });

  it('reports a closing balance that disagrees with the last row', () => {
    const { verification } = verifyBankStatement(
      statement([txn('credit', 100, 1100)], { openingBalance: 1000, closingBalance: 1200 }),
    );
    expect(verification.state).toBe('verified_with_exceptions');
    expect(verification.notes[0]).toMatch(/closing balance \(1200\.00\) does not equal the last imported row's balance \(1100\.00\)/);
  });

  it('counts rows the parser already flagged', () => {
    const { verification } = verifyBankStatement(
      statement([txn('credit', 100, 1100), txn('debit', 0, 1100, ['No amount on this row'])], {
        openingBalance: 1000,
      }),
    );
    expect(verification).toMatchObject({ state: 'verified_with_exceptions', rowsFailed: 0, rowsFlagged: 1 });
  });

  it('flags a row with no balance and keeps checking past it', () => {
    const { rows, verification } = verifyBankStatement(
      statement([txn('credit', 100, 1100), txn('debit', 100, undefined), txn('debit', 100, 900)], {
        openingBalance: 1000,
      }),
    );
    expect(rows[1].flags).toEqual(['No balance on this row']);
    expect(verification).toMatchObject({ rowsFailed: 0, rowsChecked: 2, state: 'verified_with_exceptions' });
  });

  describe('lines that were left out', () => {
    const rows = () => [txn('credit', 100, 1100), txn('debit', 100, 1000)];
    const unrecognised = [{ ref: 'row 9', text: 'stray', reason: 'unrecognised' as const }];

    it('does not count against the result when the statement totals match', () => {
      const { verification } = verifyBankStatement(
        statement(rows(), { openingBalance: 1000, statedTotalDebit: 100, statedTotalCredit: 100, skipped: unrecognised }),
      );
      expect(verification).toMatchObject({ state: 'verified', notes: [] });
    });

    it('counts against the result when there are no totals to check against', () => {
      const { verification } = verifyBankStatement(statement(rows(), { openingBalance: 1000, skipped: unrecognised }));
      expect(verification.state).toBe('verified_with_exceptions');
      expect(verification.notes).toEqual(['1 line in the file was not recognised and was left out.']);
    });

    it('always counts a transaction-like line found below the total', () => {
      const { verification } = verifyBankStatement(
        statement(rows(), {
          openingBalance: 1000,
          statedTotalDebit: 100,
          statedTotalCredit: 100,
          skipped: [{ ref: 'row 9', text: '01/04/2026 X 1.00', reason: 'after_total' }],
        }),
      );
      expect(verification.state).toBe('verified_with_exceptions');
      expect(verification.notes[0]).toMatch(/below the statement's total line/);
    });
  });

  it('does not change the rows it was given', () => {
    const input = statement([txn('credit', 100, 1100), txn('credit', 40, 1060)], { openingBalance: 1000 });
    const before = JSON.stringify(input);
    verifyBankStatement(input);
    expect(JSON.stringify(input)).toBe(before);
  });
});
