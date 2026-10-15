import { TxnSource } from '../../../bank-recon/types';
import { cleanText, normHeader } from './bank-utils';
import { detectHeader } from './table-parser';

export interface TableRoleGuess {
  role: TxnSource | 'UNKNOWN';
  /** Positive = bank statement, negative = books ledger. */
  score: number;
}

/**
 * Classify a sheet as Bank Statement vs Books Bank Ledger from its sheet name and header
 * content ("Withdrawal Amt." / "Value Date" imply bank; "Vch Type" / "Vch No." /
 * "Particulars" imply a Tally-style ledger).
 */
export function detectTableRole(grid: unknown[][], sheetName: string): TableRoleGuess {
  let score = 0;
  const name = sheetName.toLowerCase();
  if (/statement|passbook|stmt/.test(name)) score += 1;
  if (/ledger|cash\s*book|bank\s*book|books|tally|zoho|busy/.test(name)) score -= 1;

  const header = detectHeader(grid);
  if (header) {
    const cells = (grid[header.headerRow] ?? []).map((c) => normHeader(c));
    if (header.cols.vchType !== undefined || header.cols.vchNo !== undefined) score -= 2;
    if (header.cols.valueDate !== undefined) score += 2;
    if (cells.some((c) => /^(withdrawal|deposit)/.test(c))) score += 2;
    if (cells.some((c) => c === 'particulars') && header.cols.balance === undefined) score -= 1;
    if (header.cols.balance !== undefined && header.cols.vchType === undefined) score += 1;
    const pre = (grid.slice(0, header.headerRow) as unknown[][])
      .map((r) => r.map((c) => cleanText(c)).join(' '))
      .join(' ');
    if (/\bifsc\b|account\s*(no|number)|a\/c\s*no|customer\s*id/i.test(pre)) score += 1;
  }
  return { role: score >= 2 ? 'BANK' : score <= -2 ? 'BOOKS' : 'UNKNOWN', score };
}
