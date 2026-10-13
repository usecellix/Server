/** Shared date / amount / text helpers for bank-statement and books-ledger parsing. */

const MONTHS: Record<string, number> = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
};

function iso(y: number, m: number, d: number): string {
  if (m < 1 || m > 12 || d < 1 || d > 31) return '';
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return '';
  return dt.toISOString().slice(0, 10);
}

/**
 * Parse a cell into an ISO date, or '' when it is not a recognizable date.
 * Handles Excel serials, Date objects, dd/mm/yy(yy), dd-MMM-yy(yy), dd MMM yyyy, yyyy-mm-dd.
 * Indian statements are day-first, so ambiguous numeric dates are read as dd/mm.
 */
export function parseBankDate(raw: unknown): string {
  if (raw === null || raw === undefined || raw === '') return '';
  if (raw instanceof Date) {
    return Number.isNaN(raw.getTime()) ? '' : raw.toISOString().slice(0, 10);
  }
  if (typeof raw === 'number') {
    if (raw > 20000 && raw < 80000) {
      return new Date(Date.UTC(1899, 11, 30) + Math.floor(raw) * 86400000)
        .toISOString()
        .slice(0, 10);
    }
    return '';
  }
  const s = String(raw).trim();
  if (!s) return '';
  if (/^\d{5}(\.\d+)?$/.test(s)) return parseBankDate(Number(s));

  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T\s].*)?$/);
  if (m) return iso(+m[1], +m[2], +m[3]);

  m = s.match(/^(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{2}|\d{4})(?:\s.*)?$/);
  if (m) {
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    return iso(y, +m[2], +m[1]);
  }

  m = s.match(/^(\d{1,2})[-/. ]([A-Za-z]{3})[a-z]*[-/. ,]+(\d{2}|\d{4})(?:\s.*)?$/);
  if (m) {
    const mon = MONTHS[m[2].toLowerCase()];
    if (!mon) return '';
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    return iso(y, mon, +m[1]);
  }
  return '';
}

export function roundMoney(n: number): number {
  return Math.round(n * 100) / 100;
}

export interface ParsedAmount {
  value: number;
  /** 'CR' / 'DR' when the cell carried a suffix like "1,200.00 Cr". */
  suffix?: 'CR' | 'DR';
  /** True when the cell held a non-empty, parseable number. */
  present: boolean;
}

/** Parse amounts with thousands separators, ₹/Rs, parentheses negatives and Cr/Dr suffixes. */
export function parseBankAmount(raw: unknown): ParsedAmount {
  if (raw === null || raw === undefined || raw === '') return { value: 0, present: false };
  if (typeof raw === 'number') {
    return Number.isFinite(raw) ? { value: roundMoney(raw), present: true } : { value: 0, present: false };
  }
  let s = String(raw).trim();
  if (!s || s === '-' || s === '--') return { value: 0, present: false };
  let suffix: 'CR' | 'DR' | undefined;
  const sm = s.match(/\s*(cr|dr)\.?\s*$/i);
  if (sm) {
    suffix = sm[1].toUpperCase() as 'CR' | 'DR';
    s = s.slice(0, sm.index);
  }
  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1);
  }
  s = s.replace(/₹|rs\.?|inr/gi, '').replace(/,/g, '').replace(/\s+/g, '');
  if (s.startsWith('-')) {
    negative = !negative;
    s = s.slice(1);
  }
  if (!/^\d*\.?\d+$/.test(s)) return { value: 0, present: false };
  const n = Number(s);
  if (!Number.isFinite(n)) return { value: 0, present: false };
  return { value: roundMoney(negative ? -n : n), suffix, present: true };
}

export function normHeader(raw: unknown): string {
  return String(raw ?? '')
    .toLowerCase()
    .replace(/[._\-:()]/g, ' ')
    .replace(/\s*\/\s*/g, '/')
    .replace(/\s+/g, ' ')
    .trim();
}

export function cleanText(raw: unknown): string {
  return String(raw ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

export function daysDiff(a: string, b: string): number {
  return Math.round((Date.parse(b) - Date.parse(a)) / 86400000);
}

export function addDays(isoDate: string, days: number): string {
  return new Date(Date.parse(isoDate) + days * 86400000).toISOString().slice(0, 10);
}
