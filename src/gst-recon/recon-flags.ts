/**
 * P2.3 + P4.3 — per-row flags added to the GST recon output.
 *  - GSTIN check (B10 checksum validator): flags, never corrects.
 *  - Normalised invoice key (shared B10 normaliser): why "INV/24-25/0233" matched "233".
 *  - Sec 16(4) time-limit status, from the rule table (CA to verify) — a flag, not a verdict.
 */
import { GSTIN_STATUS_LABEL, normalizeInvoiceKey, validateGstin } from '../excel-ai/utils/india-validators.util';
import { getRuleTable, ruleCitation } from '../rules/rule-tables';

export function gstinCheckLabel(gstin: unknown): string | null {
  const raw = String(gstin ?? '').trim();
  if (!raw) return null;
  return GSTIN_STATUS_LABEL[validateGstin(raw).status];
}

export const invoiceKey = (invoiceNumber: unknown): string | null => {
  const key = normalizeInvoiceKey(invoiceNumber);
  return key || null;
};

function parseInvoiceDate(v: unknown): number | null {
  if (typeof v === 'number' && v > 20_000 && v < 80_000) return Math.round((v - 25_569) * 86_400_000);
  const t = String(v ?? '').trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(t);
  if (m) return Date.UTC(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
  return null;
}

export type Sec164Status = 'Within' | 'Near' | 'Beyond';

/** ITC time limit for the invoice's FY: 30 Nov of the following FY (rule table), judged on `asOn`. */
export function sec164Status(invoiceDate: unknown, asOn: Date = new Date()): { status: Sec164Status; deadline: string } | null {
  const ms = parseInvoiceDate(invoiceDate);
  if (ms === null) return null;
  const d = new Date(ms);
  const fyStart = d.getUTCMonth() >= 3 ? d.getUTCFullYear() : d.getUTCFullYear() - 1;
  const rule = getRuleTable('gst-time-limits').data;
  const deadline = Date.UTC(fyStart + 1, rule.itcDeadlineMonth - 1, rule.itcDeadlineDay);
  const today = Date.UTC(asOn.getUTCFullYear(), asOn.getUTCMonth(), asOn.getUTCDate());
  const daysLeft = Math.round((deadline - today) / 86_400_000);
  const status: Sec164Status = daysLeft < 0 ? 'Beyond' : daysLeft <= rule.nearDays ? 'Near' : 'Within';
  const dd = new Date(deadline);
  return { status, deadline: `${String(dd.getUTCDate()).padStart(2, '0')}-${String(dd.getUTCMonth() + 1).padStart(2, '0')}-${dd.getUTCFullYear()}` };
}

export const SEC_16_4_CITATION = (): string => ruleCitation('gst-time-limits');
