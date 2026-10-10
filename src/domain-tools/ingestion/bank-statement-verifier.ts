import { BankStatementParseResult, NormalizedBankStatementRow } from './bank-statement-parser';
import { formatPaise, paiseToAmount } from './statement-values';

/**
 * Arithmetic checks on a parsed statement (Root/ATTACHMENT_EXTRACTION_PLAN.md §6.4).
 *
 * The bank prints a balance after every transaction, so each row can be checked
 * against the one before it: previous balance + credit - debit must equal this
 * row's balance. That check is what lets an import be called verified without a
 * person or a model reading it.
 */

export type BankStatementVerificationState =
  /** Every row ties and nothing was left out. */
  | 'verified'
  /** Imported, but the listed rows or totals need a look. */
  | 'verified_with_exceptions'
  /** The statement has no balance column, so there was nothing to check against. */
  | 'unverified';

export interface BankStatementVerification {
  state: BankStatementVerificationState;
  /** Plain sentences explaining anything short of `verified`, for the result card. */
  notes: string[];
  /** The order the statement lists transactions in. Rows keep that order. */
  order: 'oldest_first' | 'newest_first';
  rowsChecked: number;
  rowsFailed: number;
  rowsFlagged: number;
  /** Rows whose debit/credit side was set from the balance because the column position disagreed with it. */
  directionFixes: number;
  openingBalance?: number;
  openingBalanceSource: 'stated' | 'derived' | 'none';
  closingBalance?: number;
  totalDebit: number;
  totalCredit: number;
  totalsCheck: 'match' | 'mismatch' | 'not_stated';
}

export interface VerifiedBankStatement {
  /** Same rows in the same order, with flags added and any direction fix applied. */
  rows: NormalizedBankStatementRow[];
  verification: BankStatementVerification;
}

const toPaise = (amount: number): number => Math.round(amount * 100);

/** Signed movement in paise: credits positive, debits negative. */
const movement = (row: NormalizedBankStatementRow): number =>
  row.type === 'credit' ? toPaise(row.amount) : -toPaise(row.amount);

/**
 * Statements run oldest-first or newest-first. Whichever direction makes more
 * consecutive balances differ by exactly the row's own amount is the real one.
 */
function detectOrder(rows: NormalizedBankStatementRow[]): BankStatementVerification['order'] {
  let forward = 0;
  let backward = 0;
  for (let i = 0; i < rows.length; i++) {
    const balance = rows[i].balance;
    if (balance === undefined) continue;
    const size = toPaise(rows[i].amount);
    const before = rows[i - 1]?.balance;
    const after = rows[i + 1]?.balance;
    if (before !== undefined && Math.abs(toPaise(balance) - toPaise(before)) === size) forward++;
    if (after !== undefined && Math.abs(toPaise(balance) - toPaise(after)) === size) backward++;
  }
  return backward > forward ? 'newest_first' : 'oldest_first';
}

export function verifyBankStatement(parsed: BankStatementParseResult): VerifiedBankStatement {
  const rows = parsed.rows.map((row) => ({ ...row, flags: [...row.flags] }));
  const notes: string[] = [];

  const order = parsed.hasBalanceColumn ? detectOrder(rows) : 'oldest_first';
  const chronological = order === 'newest_first' ? [...rows].reverse() : rows;

  const statedOpening =
    parsed.openingBalance !== undefined ? toPaise(parsed.openingBalance) : undefined;
  let previous = statedOpening;
  let rowsChecked = 0;
  let rowsFailed = 0;
  let directionFixes = 0;

  if (parsed.hasBalanceColumn) {
    for (const row of chronological) {
      if (row.balance === undefined) {
        row.flags.push('No balance on this row');
        if (previous !== undefined) previous += movement(row);
        continue;
      }
      const balance = toPaise(row.balance);
      if (previous === undefined) {
        // First row of a statement that prints no opening balance: nothing to check it against.
        previous = balance;
        continue;
      }
      rowsChecked++;
      if (previous + movement(row) === balance) {
        previous = balance;
        continue;
      }
      // The printed balance is authoritative on direction. If the row ties
      // with the opposite sign, the amount sat under the wrong heading.
      if (row.amount !== 0 && previous - movement(row) === balance) {
        row.type = row.type === 'credit' ? 'debit' : 'credit';
        directionFixes++;
        previous = balance;
        continue;
      }
      rowsFailed++;
      row.flags.push(
        `Balance does not tie: expected ${formatPaise(previous + movement(row))}, statement shows ${formatPaise(balance)}`,
      );
      // Resynchronise so one bad row is reported once, not on every row after it.
      previous = balance;
    }
  }

  let totalDebit = 0;
  let totalCredit = 0;
  for (const row of rows) {
    if (row.type === 'credit') totalCredit += toPaise(row.amount);
    else totalDebit += toPaise(row.amount);
  }

  let totalsCheck: BankStatementVerification['totalsCheck'] = 'not_stated';
  if (parsed.statedTotalDebit !== undefined || parsed.statedTotalCredit !== undefined) {
    const debitOk =
      parsed.statedTotalDebit === undefined || toPaise(parsed.statedTotalDebit) === totalDebit;
    const creditOk =
      parsed.statedTotalCredit === undefined || toPaise(parsed.statedTotalCredit) === totalCredit;
    totalsCheck = debitOk && creditOk ? 'match' : 'mismatch';
    if (totalsCheck === 'mismatch') {
      notes.push(
        `The statement's own totals (withdrawals ${parsed.statedTotalDebit ?? '-'}, deposits ${parsed.statedTotalCredit ?? '-'}) ` +
          `do not equal the imported rows (withdrawals ${formatPaise(totalDebit)}, deposits ${formatPaise(totalCredit)}).`,
      );
    }
  }

  const first = chronological[0];
  const last = chronological[chronological.length - 1];
  let openingBalance: number | undefined;
  let openingBalanceSource: BankStatementVerification['openingBalanceSource'] = 'none';
  if (statedOpening !== undefined) {
    openingBalance = statedOpening;
    openingBalanceSource = 'stated';
  } else if (first?.balance !== undefined) {
    openingBalance = toPaise(first.balance) - movement(first);
    openingBalanceSource = 'derived';
  }
  const closingBalance = last?.balance !== undefined ? toPaise(last.balance) : undefined;

  let closingMismatch = false;
  if (parsed.closingBalance !== undefined && closingBalance !== undefined) {
    closingMismatch = toPaise(parsed.closingBalance) !== closingBalance;
    if (closingMismatch) {
      notes.push(
        `The statement's closing balance (${formatPaise(toPaise(parsed.closingBalance))}) ` +
          `does not equal the last imported row's balance (${formatPaise(closingBalance)}).`,
      );
    }
  }

  const rowsFlagged = rows.filter((row) => row.flags.length > 0).length;
  if (rowsFlagged > 0) {
    notes.push(
      `${rowsFlagged} row${rowsFlagged === 1 ? '' : 's'} could not be fully checked. See the Flag column.`,
    );
  }

  const afterTotal = parsed.skipped.filter((row) => row.reason === 'after_total').length;
  const unrecognised = parsed.skipped.length - afterTotal;
  if (afterTotal > 0) {
    notes.push(
      afterTotal === 1
        ? "1 line below the statement's total line looked like a transaction and was not imported."
        : `${afterTotal} lines below the statement's total line looked like transactions and were not imported.`,
    );
  }
  // Lines that are not transactions cannot be missing money when the bank's
  // own totals agree with what was imported, so they only count against the
  // result when that cross-check is unavailable.
  const unrecognisedMatters = unrecognised > 0 && totalsCheck !== 'match';
  if (unrecognisedMatters) {
    notes.push(
      unrecognised === 1
        ? '1 line in the file was not recognised and was left out.'
        : `${unrecognised} lines in the file were not recognised and were left out.`,
    );
  }

  let state: BankStatementVerificationState;
  if (!parsed.hasBalanceColumn) {
    state = 'unverified';
    notes.unshift(
      totalsCheck === 'match'
        ? 'The statement has no balance column, so rows could not be checked one by one. The totals do match the statement.'
        : 'The statement has no balance column, so the rows could not be checked.',
    );
  } else if (
    rowsFailed === 0 &&
    rowsFlagged === 0 &&
    totalsCheck !== 'mismatch' &&
    !closingMismatch &&
    afterTotal === 0 &&
    !unrecognisedMatters
  ) {
    state = 'verified';
  } else {
    state = 'verified_with_exceptions';
  }

  return {
    rows,
    verification: {
      state,
      notes,
      order,
      rowsChecked,
      rowsFailed,
      rowsFlagged,
      directionFixes,
      ...(openingBalance !== undefined ? { openingBalance: paiseToAmount(openingBalance) } : {}),
      openingBalanceSource,
      ...(closingBalance !== undefined ? { closingBalance: paiseToAmount(closingBalance) } : {}),
      totalDebit: paiseToAmount(totalDebit),
      totalCredit: paiseToAmount(totalCredit),
      totalsCheck,
    },
  };
}
