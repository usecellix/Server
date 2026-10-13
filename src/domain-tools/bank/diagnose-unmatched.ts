import {
  BankMatchSettings,
  BankReason,
  BankTxn,
  BooksReason,
  UnmatchedItem,
} from '../../bank-recon/types';
import { daysDiff } from '../ingestion/bank/bank-utils';

export interface DiagnoseInput {
  allBooks: BankTxn[];
  allBank: BankTxn[];
  freeBooks: BankTxn[];
  freeBank: BankTxn[];
  usedBooks: Set<string>;
  usedBank: Set<string>;
  settings: BankMatchSettings;
  periodEnd?: string;
  refMatch: (a: BankTxn, b: BankTxn) => boolean;
  refConflict: (a: BankTxn, b: BankTxn) => boolean;
  /** bank − books, days. */
  gapOf: (books: BankTxn, bank: BankTxn) => number;
}

const STALE_DAYS = 90;
const FAR_WINDOW_DAYS = 90;
const ELECTRONIC = new Set(['NEFT', 'RTGS', 'IMPS', 'UPI']);

/** Electronic when the channel says so, or when it carries a UTR-style reference but no cheque number. */
const isElectronic = (t: BankTxn): boolean =>
  (t.channel ? ELECTRONIC.has(t.channel) : false) || (!t.channel && !!t.referenceNo && !t.instrumentNo);

const inr = (n: number) => `₹${n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

function dupKey(t: BankTxn): string {
  return [t.txnDate, t.direction, Math.round(t.amount * 100), t.narration.toUpperCase().trim(), t.instrumentNo ?? ''].join('|');
}

/** Bank-originated item classification from narration keywords (plan §6 Pass 6). */
export function classifyBankOnly(t: BankTxn): { reason: BankReason; explanation: string } {
  const n = t.narration.toUpperCase();
  if (/(CHQ|CHEQUE|CLG|INW|OUTW|ECS|NACH).*(RET|RTN|RETURN|BOUNCE|UNPAID|DISHON)|(RET|RETURN|BOUNCE|UNPAID|DISHON).*(CHQ|CHEQUE|CHG|CHARGE)/.test(n)) {
    return {
      reason: 'cheque_bounce_not_booked',
      explanation: 'Cheque / mandate return or bounce charge on the statement — post the reversal / charge in books.',
    };
  }
  if (t.direction === 'OUT' && /\bTDS\b|TAX DEDUCTED/.test(n)) {
    return { reason: 'tds_on_interest_not_booked', explanation: 'TDS deducted by the bank — book as TDS receivable.' };
  }
  if (t.direction === 'IN' && /\b(INT\.?\s?(PD|PAID|CR|CREDIT)?|INTEREST|INT\s?PAID)\b/.test(n)) {
    return { reason: 'interest_credited_not_booked', explanation: 'Interest credited by the bank — book as interest income.' };
  }
  if (t.direction === 'OUT' && /\b(GST|IGST|CGST|SGST)\b/.test(n) && /CHG|CHARGE|FEE|SERVICE|SMS/.test(n)) {
    return { reason: 'bank_charges_not_booked', explanation: 'GST on bank charges — book along with the charges.' };
  }
  if (
    t.direction === 'OUT' &&
    /\b(CHG|CHGS|CHRG|CHARGES?|FEES?|PENALTY|AMC|MIN(IMUM)?\s?BAL|SMS|PROCESSING|SERVICE\s?CH)\b/.test(n)
  ) {
    return { reason: 'bank_charges_not_booked', explanation: 'Bank charge not yet booked in the ledger.' };
  }
  if (
    t.direction === 'OUT' &&
    (t.channel === 'ECS' || t.channel === 'NACH' ||
      /\b(EMI|SIP|LIC|INSURANCE|LOAN|MANDATE|AUTO\s?DEBIT|STANDING\s?INSTR|CHALLAN|E-?PAYMENT|GST\s?PAYMENT|GST\s?PMT|PMT-?06|TAX)\b/.test(n))
  ) {
    return { reason: 'direct_debit_not_booked', explanation: 'Auto-debit / direct payment on the statement with no ledger entry.' };
  }
  if (t.direction === 'IN' && /\b(RET|RETURN(ED)?|REVERSAL|REFUND|FAILED)\b/.test(n)) {
    return {
      reason: 'return_credit_not_booked',
      explanation: 'A payment was returned / refunded by the bank — reverse or re-post the original payment in books.',
    };
  }
  if (t.direction === 'IN' && /UNKNOWN|UNIDENT|NOT\s?KNOWN/.test(n)) {
    return {
      reason: 'unidentified_transaction',
      explanation: 'Credit from an unidentified sender — needs CA review before posting.',
    };
  }
  if (t.direction === 'IN') {
    return {
      reason: 'direct_credit_not_booked',
      explanation: 'Credit in the bank with no receipt entry in the books — identify the party and post the receipt.',
    };
  }
  return {
    reason: 'unidentified_transaction',
    explanation: 'Debit in the bank with no matching payment entry — needs CA review.',
  };
}


/**
 * Key for "the same bank debit/credit posted twice". Normally the same reference; a card
 * (POS) machine double-charge carries a different RRN each time, so for POS the digit tokens
 * (card number, RRN) are dropped and the merchant text decides.
 */
function bankDupKey(k: BankTxn): string | null {
  const base = [k.txnDate, k.direction, Math.round(k.amount * 100)].join('|');
  if (k.channel === 'POS') {
    const merchant = k.narration.toUpperCase().split(/\s+/).filter((w) => !/\d/.test(w)).join(' ');
    return merchant ? base + '|POS|' + merchant : null;
  }
  if (!k.referenceNo && !k.instrumentNo) return null;
  return base + '|' + (k.referenceNo ?? '') + '|' + (k.instrumentNo ?? '');
}

export function diagnoseUnmatched(input: DiagnoseInput): {
  unmatchedBooks: UnmatchedItem<BooksReason>[];
  unmatchedBank: UnmatchedItem<BankReason>[];
} {
  const { settings, periodEnd } = input;

  // Duplicate groups in books / bank
  const booksGroups = new Map<string, BankTxn[]>();
  for (const b of input.allBooks) (booksGroups.get(dupKey(b)) ?? booksGroups.set(dupKey(b), []).get(dupKey(b))!).push(b);

  const unmatchedBooks: UnmatchedItem<BooksReason>[] = input.freeBooks.map((b) => {
    const mk = (reason: BooksReason, explanation: string): UnmatchedItem<BooksReason> => ({ txn: b, reason, explanation });

    const twins = booksGroups.get(dupKey(b)) ?? [];
    if (twins.length > 1) {
      const idx = twins.findIndex((t) => t.id === b.id);
      const otherMatched = twins.some((t) => t.id !== b.id && input.usedBooks.has(t.id));
      if (otherMatched || idx > 0) {
        return mk('duplicate_in_books', 'Identical entry booked more than once — likely a duplicate posting.');
      }
    }

    // Same reference, different amount
    const refHit = input.freeBank.find(
      (k) => k.direction === b.direction && input.refMatch(b, k) && Math.round(k.amount * 100) !== Math.round(b.amount * 100),
    );
    if (refHit) {
      return mk(
        'amount_mismatch',
        `Same cheque/reference appears in the bank for ${inr(refHit.amount)} but is ${inr(b.amount)} in books (difference ${inr(Math.abs(refHit.amount - b.amount))}).`,
      );
    }

    // Same reference/amount but outside the clearing window
    const far = input.freeBank.find((k) => {
      if (k.direction !== b.direction || Math.round(k.amount * 100) !== Math.round(b.amount * 100)) return false;
      if (input.refConflict(b, k)) return false;
      const g = input.gapOf(b, k);
      return (g > settings.clearingWindowDays && g <= FAR_WINDOW_DAYS) || (g < -settings.earlyWindowDays && g >= -FAR_WINDOW_DAYS);
    });
    if (far) {
      return mk(
        'date_mismatch_outside_window',
        `A bank entry for ${inr(far.amount)} exists on ${far.txnDate}, ${Math.abs(input.gapOf(b, far))} days from the books date — check the date posted.`,
      );
    }

    if (periodEnd && b.txnDate > periodEnd) {
      return mk('post_dated_cheque', `Dated ${b.txnDate}, after the period end ${periodEnd} — outside this reconciliation.`);
    }

    const electronic = isElectronic(b);
    const age = periodEnd ? daysDiff(b.txnDate, periodEnd) : 0;

    if (b.direction === 'OUT') {
      if (electronic && age > settings.tightWindowDays) {
        return mk('genuinely_missing', `Electronic payment (${b.channel ?? 'by UTR'}) booked ${age} days before period end but absent from the bank statement.`);
      }
      if (age > STALE_DAYS) {
        return mk('stale_cheque', `Issued ${age} days before period end and still uncleared — cheques are valid for 3 months; consider reversing.`);
      }
      return mk('cheque_issued_not_cleared', 'Payment booked but not yet presented / debited by the bank at period end.');
    }
    // IN
    if (electronic && age > settings.tightWindowDays) {
      return mk('genuinely_missing', `Electronic receipt (${b.channel ?? 'by UTR'}) booked ${age} days before period end but absent from the bank statement.`);
    }
    if (age > 30) {
      return mk('genuinely_missing', `Receipt booked ${age} days before period end and still not credited by the bank.`);
    }
    return mk('deposit_in_transit', 'Receipt booked but not yet credited by the bank at period end.');
  });

  // Bank side
  const bankGroups = new Map<string, BankTxn[]>();
  for (const k of input.allBank) {
    const key = bankDupKey(k);
    if (!key) continue;
    (bankGroups.get(key) ?? bankGroups.set(key, []).get(key)!).push(k);
  }

  const unmatchedBank: UnmatchedItem<BankReason>[] = input.freeBank.map((k) => {
    const key = bankDupKey(k);
    const twins = key ? (bankGroups.get(key) ?? []) : [];
    const otherTwinMatched = twins.some((t) => t.id !== k.id && input.usedBank.has(t.id));
    if (twins.length > 1 && (otherTwinMatched || twins.findIndex((t) => t.id === k.id) > 0)) {
      return { txn: k, reason: 'possible_duplicate_in_bank' as const, explanation: 'The same debit/credit appears twice on the statement (same date, amount and reference/merchant) — check for a double charge.' };
    }
    const c = classifyBankOnly(k);
    // Hint when a books row looks like the same transaction but sits outside the window.
    const near = input.freeBooks.find(
      (b) => b.direction === k.direction && Math.round(b.amount * 100) === Math.round(k.amount * 100) && !input.refConflict(b, k),
    );
    const refBooks = input.freeBooks.find((b) => b.direction === k.direction && input.refMatch(b, k));
    const refHint = refBooks && !near ? ` Same reference as books entry dated ${refBooks.txnDate} (row ${refBooks.sourceLocation.row}) for a different amount.` : '';
    const hint = refHint || (near ? ` Possibly the books entry dated ${near.txnDate} (row ${near.sourceLocation.row}).` : "");
    return { txn: k, reason: c.reason, explanation: c.explanation + hint };
  });

  return { unmatchedBooks, unmatchedBank };
}

