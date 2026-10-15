import { BankTxn, ParsedTable, ValidationIssue, ValidationResult } from '../../../bank-recon/types';
import { addDays } from './bank-utils';

/** Validation gate that runs before matching — see plan §5.5. */

const TOL = 0.011;

const signed = (t: BankTxn): number => (t.direction === 'IN' ? t.amount : -t.amount);

export interface ValidateOptions {
  label: string; // "Bank statement" | "Books ledger"
  periodStart?: string;
  periodEnd?: string;
  bufferDays?: number;
  /** Bank only — account number the user / ledger mapped this run to. */
  accountNumber?: string;
  /** User-supplied balances override what the sheet printed. */
  openingBalance?: number | null;
  closingBalance?: number | null;
}

/**
 * Drops rows repeated by overlapping / merged page ranges. Only rows that carry a running
 * balance are eligible — two genuinely identical payments on one day have different
 * balances, so identical-including-balance rows can only be a repeated page.
 */
export function dedupeRepeatedRows(txns: BankTxn[]): { txns: BankTxn[]; removed: BankTxn[] } {
  const seen = new Set<string>();
  const kept: BankTxn[] = [];
  const removed: BankTxn[] = [];
  for (const t of txns) {
    if (t.runningBalance === undefined) {
      kept.push(t);
      continue;
    }
    const key = [t.txnDate, t.direction, t.amount, t.runningBalance, t.narration.toUpperCase()].join('|');
    if (seen.has(key)) removed.push(t);
    else {
      seen.add(key);
      kept.push(t);
    }
  }
  return { txns: kept, removed };
}

export function validateParsedTable(parsed: ParsedTable, opts: ValidateOptions): ValidationResult {
  const errors: ValidationIssue[] = [];
  const warnings: ValidationIssue[] = [];
  const { txns } = parsed;
  const opening = opts.openingBalance ?? parsed.openingBalance;
  const closing = opts.closingBalance ?? parsed.closingBalance;

  // 1. Balance tie-out: opening + in - out = closing
  if (opening !== null && closing !== null && txns.length) {
    const expected = opening + txns.reduce((s, t) => s + signed(t), 0);
    const diff = Math.round((expected - closing) * 100) / 100;
    if (Math.abs(diff) > TOL) {
      errors.push({
        code: 'BALANCE_TIE_OUT',
        message:
          `${opts.label}: opening ${fmt(opening)} + receipts − payments gives ${fmt(expected)}, ` +
          `but the closing balance is ${fmt(closing)} (off by ${fmt(diff)}). ` +
          `A row may be missing, duplicated or on the wrong side.`,
      });
    }
  }

  // 2. Running balance continuity — localizes the first dropped / duplicated / mis-signed row.
  const withBal = txns.filter((t) => t.runningBalance !== undefined);
  if (withBal.length >= 2 && withBal.length >= txns.length * 0.5) {
    let prev: number | undefined = opening ?? undefined;
    let breaks = 0;
    let firstBreak: BankTxn | undefined;
    for (const t of txns) {
      if (t.runningBalance === undefined) {
        prev = undefined;
        continue;
      }
      if (prev !== undefined) {
        const expected = prev + signed(t);
        if (Math.abs(expected - t.runningBalance) > TOL) {
          breaks += 1;
          firstBreak = firstBreak ?? t;
        }
      }
      prev = t.runningBalance;
    }
    if (breaks > 0 && firstBreak) {
      errors.push({
        code: 'BALANCE_CONTINUITY',
        row: firstBreak.sourceLocation.row,
        message:
          `${opts.label}: running balance stops adding up at row ${firstBreak.sourceLocation.row} ` +
          `(${breaks} break${breaks > 1 ? 's' : ''} in total) — check that row and the one above it.`,
      });
    }
  }

  // 3. Date sanity
  if (opts.periodStart && opts.periodEnd) {
    const buf = opts.bufferDays ?? 15;
    const lo = addDays(opts.periodStart, -buf);
    const hi = addDays(opts.periodEnd, buf);
    const outliers = txns.filter((t) => t.txnDate < lo || t.txnDate > hi);
    if (outliers.length) {
      warnings.push({
        code: 'DATE_OUTLIER',
        row: outliers[0].sourceLocation.row,
        message: `${opts.label}: ${outliers.length} row(s) fall outside ${opts.periodStart} to ${opts.periodEnd} (±${buf} days), first at row ${outliers[0].sourceLocation.row}.`,
      });
    }
  }

  // Low-confidence rows
  const low = txns.filter((t) => t.parseConfidence < 0.9);
  if (low.length) {
    warnings.push({
      code: 'LOW_CONFIDENCE_ROWS',
      row: low[0].sourceLocation.row,
      message: `${opts.label}: ${low.length} row(s) were read with lower confidence (direction inferred or date borrowed) — please verify.`,
    });
  }

  // 4. Account match (bank statement only)
  if (opts.accountNumber) {
    const wanted = opts.accountNumber.replace(/\D/g, '');
    if (wanted.length >= 4) {
      const last4 = wanted.slice(-4);
      const found = parsed.preambleText
        .join(' ')
        .match(/[Xx*]*\d[\dXx*]{7,}/g)
        ?.map((s) => s.replace(/\D/g, ''))
        .filter((s) => s.length >= 4);
      if (found?.length && !found.some((n) => n.endsWith(last4))) {
        errors.push({
          code: 'ACCOUNT_MISMATCH',
          message: `${opts.label}: the account number on the statement header doesn't end in ${last4} — is this the right account?`,
        });
      }
    }
  }

  warnings.push(...parsed.warnings.slice(0, 5).map((m) => ({ code: 'PARSE_NOTE', message: m })));
  return { ok: errors.length === 0, errors, warnings };
}

function fmt(n: number): string {
  return n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
