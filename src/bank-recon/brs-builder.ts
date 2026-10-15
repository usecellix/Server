import { BankMatchResult, BankTxn, BrsLine, BrsResult } from './types';

const r2 = (n: number) => Math.round(n * 100) / 100;
const signed = (t: BankTxn) => (t.direction === 'IN' ? t.amount : -t.amount);

/** Balance as at `asAt`: last printed running balance on/before that date, else opening + movements. */
export function balanceAsAt(
  txns: BankTxn[],
  opening: number | null,
  asAt: string,
): number | null {
  const upTo = txns.filter((t) => t.txnDate <= asAt);
  for (let i = upTo.length - 1; i >= 0; i--) {
    if (upTo[i].runningBalance !== undefined) return upTo[i].runningBalance!;
  }
  if (opening === null) return null;
  return r2(opening + upTo.reduce((s, t) => s + signed(t), 0));
}

export interface BuildBrsInput {
  result: BankMatchResult;
  periodEnd: string;
  booksBalance: number | null;
  bankBalance: number | null;
  /** Extra context for the not-tied-out note. */
  validationWarnings?: string[];
}

/**
 * Standard two-sided BRS as at the period end:
 *   Books balance  + bank-only receipts − bank-only payments   = Adjusted books
 *   Bank balance   + books-only receipts − books-only payments = Adjusted bank
 * Cheques issued after the period end but cleared before it (and vice-versa) fall out of
 * the match results automatically because each row is compared against `periodEnd`.
 */
export function buildBrs(input: BuildBrsInput): BrsResult {
  const { result, periodEnd } = input;
  let bankOnlyIn = 0;
  let bankOnlyOut = 0;
  let booksOnlyIn = 0;
  let booksOnlyOut = 0;

  for (const u of result.unmatchedBooks) {
    if (u.txn.txnDate > periodEnd) continue;
    if (u.txn.direction === 'IN') booksOnlyIn += u.txn.amount;
    else booksOnlyOut += u.txn.amount;
  }
  for (const u of result.unmatchedBank) {
    if (u.txn.txnDate > periodEnd) continue;
    if (u.txn.direction === 'IN') bankOnlyIn += u.txn.amount;
    else bankOnlyOut += u.txn.amount;
  }

  for (const m of result.matches) {
    if (m.kind === 'REVERSAL') {
      for (const t of [...m.books, ...m.bank]) {
        if (t.txnDate > periodEnd) continue;
        const side = m.side === 'BOOKS' ? 'books' : 'bank';
        if (side === 'books') t.direction === 'IN' ? (booksOnlyIn += t.amount) : (booksOnlyOut += t.amount);
        else t.direction === 'IN' ? (bankOnlyIn += t.amount) : (bankOnlyOut += t.amount);
      }
      continue;
    }
    const dir = (m.books[0] ?? m.bank[0]).direction;
    const booksAmt = m.books.filter((b) => b.txnDate <= periodEnd).reduce((s, b) => s + b.amount, 0);
    const bankAmt = m.bank.filter((k) => k.txnDate <= periodEnd).reduce((s, k) => s + k.amount, 0);
    const d = r2(booksAmt - bankAmt);
    if (d > 0) dir === 'IN' ? (booksOnlyIn += d) : (booksOnlyOut += d);
    else if (d < 0) dir === 'IN' ? (bankOnlyIn += -d) : (bankOnlyOut += -d);
  }

  bankOnlyIn = r2(bankOnlyIn);
  bankOnlyOut = r2(bankOnlyOut);
  booksOnlyIn = r2(booksOnlyIn);
  booksOnlyOut = r2(booksOnlyOut);

  const { booksBalance, bankBalance } = input;
  const adjustedBooks = booksBalance === null ? null : r2(booksBalance + bankOnlyIn - bankOnlyOut);
  const adjustedBank = bankBalance === null ? null : r2(bankBalance + booksOnlyIn - booksOnlyOut);
  const difference = adjustedBooks === null || adjustedBank === null ? null : r2(adjustedBooks - adjustedBank);
  const tiedOut = difference !== null && Math.abs(difference) < 0.01;

  const lines: BrsLine[] = [
    { label: 'Balance as per Books', amount: booksBalance, kind: 'balance' },
    { label: 'Add: Credits in bank not yet in books (interest, direct receipts)', amount: bankOnlyIn, kind: 'add' },
    { label: 'Less: Debits in bank not yet in books (charges, direct debits)', amount: bankOnlyOut, kind: 'less' },
    { label: 'Adjusted Balance as per Books', amount: adjustedBooks, kind: 'total' },
    { label: 'Balance as per Bank Statement', amount: bankBalance, kind: 'balance' },
    { label: 'Add: Deposits recorded in books, not yet credited by bank', amount: booksOnlyIn, kind: 'add' },
    { label: 'Less: Cheques / payments issued, not yet presented', amount: booksOnlyOut, kind: 'less' },
    { label: 'Adjusted Balance as per Bank', amount: adjustedBank, kind: 'total' },
    { label: 'Difference (must be 0.00)', amount: difference, kind: 'difference' },
  ];

  const notes: string[] = [];
  if (booksBalance === null || bankBalance === null) {
    notes.push(
      `${booksBalance === null ? 'Books' : 'Bank'} balance could not be read from the sheet — add an opening balance row (or tell me the balance) to tie the BRS out. Reconciling items below are still complete.`,
    );
  } else if (!tiedOut) {
    notes.push(
      `The BRS does not tie out — residual difference ${difference!.toLocaleString('en-IN', { minimumFractionDigits: 2 })}. This is NOT a reconciled result.`,
      'Likely causes: rows that were not parsed, a balance break in the statement/ledger, items carried over from the previous period, or ambiguous groups.',
    );
    if (input.validationWarnings?.length) notes.push(...input.validationWarnings);
  }

  return {
    periodEnd,
    booksBalance,
    bankBalance,
    bankOnlyIn,
    bankOnlyOut,
    booksOnlyIn,
    booksOnlyOut,
    adjustedBooks,
    adjustedBank,
    difference,
    tiedOut,
    lines,
    notes,
  };
}
