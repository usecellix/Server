/**
 * Date and amount reading for statement ingestion. Pure functions, no I/O.
 *
 * Money is held in paise (integers) everywhere inside the ingestion code so the
 * running-balance arithmetic is exact. Convert with `paiseToAmount` only at the
 * edge.
 */

export interface DateParts {
  y: number;
  m: number;
  d: number;
}

/** Which way round a numeric date such as 04/05/2026 is read. */
export type NumericDateOrder = 'dmy' | 'mdy';

const MONTHS: Record<string, number> = {
  jan: 1, january: 1,
  feb: 2, february: 2,
  mar: 3, march: 3,
  apr: 4, april: 4,
  may: 5,
  jun: 6, june: 6,
  jul: 7, july: 7,
  aug: 8, august: 8,
  sep: 9, sept: 9, september: 9,
  oct: 10, october: 10,
  nov: 11, november: 11,
  dec: 12, december: 12,
};

const TIME_SUFFIX_RE = /[\sT]+\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?\s*(?:[ap]\.?m\.?)?\s*$/i;
const ISO_RE = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/;
const DAY_MONTHNAME_RE = /^(\d{1,2})[-/.\s]*([a-z]{3,9})[-/.,\s]*(\d{2}|\d{4})$/i;
const MONTHNAME_DAY_RE = /^([a-z]{3,9})[-/.\s]*(\d{1,2})[-/.,\s]+(\d{2}|\d{4})$/i;
const NUMERIC_RE = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/;

function fullYear(raw: string): number {
  const n = Number(raw);
  return raw.length === 2 ? 2000 + n : n;
}

function daysInMonth(y: number, m: number): number {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

function validDate(y: number, m: number, d: number): DateParts | null {
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;
  if (y < 1990 || y > 2100) return null;
  if (m < 1 || m > 12) return null;
  if (d < 1 || d > daysInMonth(y, m)) return null;
  return { y, m, d };
}

/**
 * Reads one statement date. Returns null when the text is not a date.
 * `order` only matters for all-numeric dates; named months and ISO dates are
 * unambiguous.
 */
export function parseStatementDate(
  text: string,
  order: NumericDateOrder = 'dmy',
): DateParts | null {
  const cleaned = String(text ?? '').trim().replace(TIME_SUFFIX_RE, '').trim();
  if (!cleaned) return null;

  let match = ISO_RE.exec(cleaned);
  if (match) return validDate(Number(match[1]), Number(match[2]), Number(match[3]));

  match = DAY_MONTHNAME_RE.exec(cleaned);
  if (match) {
    const month = MONTHS[match[2].toLowerCase()];
    return month ? validDate(fullYear(match[3]), month, Number(match[1])) : null;
  }

  match = MONTHNAME_DAY_RE.exec(cleaned);
  if (match) {
    const month = MONTHS[match[1].toLowerCase()];
    return month ? validDate(fullYear(match[3]), month, Number(match[2])) : null;
  }

  match = NUMERIC_RE.exec(cleaned);
  if (match) {
    const first = Number(match[1]);
    const second = Number(match[2]);
    const year = fullYear(match[3]);
    return order === 'dmy' ? validDate(year, second, first) : validDate(year, first, second);
  }

  return null;
}

/** True when the text reads as a date under either numeric order. */
export function looksLikeDate(text: string): boolean {
  return parseStatementDate(text, 'dmy') !== null || parseStatementDate(text, 'mdy') !== null;
}

export function toIsoDate(parts: DateParts): string {
  const mm = String(parts.m).padStart(2, '0');
  const dd = String(parts.d).padStart(2, '0');
  return `${parts.y}-${mm}-${dd}`;
}

/** Days since the Unix epoch, for ordering dates without a Date object. */
export function dateOrdinal(parts: DateParts): number {
  return Math.round(Date.UTC(parts.y, parts.m - 1, parts.d) / 86_400_000);
}

/** Excel's 1900-system serial for an ISO date (valid from 1 March 1900 on). */
export function isoDateToExcelSerial(iso: string): number | null {
  const match = ISO_RE.exec(iso);
  if (!match) return null;
  const parts = validDate(Number(match[1]), Number(match[2]), Number(match[3]));
  return parts ? dateOrdinal(parts) + 25_569 : null;
}

/**
 * Picks the numeric date order for a whole column. A value over 12 in either
 * slot settles it. When every date is ambiguous, the reading that keeps the
 * column in order wins, and a tie falls back to day-first.
 */
export function inferNumericDateOrder(texts: string[]): NumericDateOrder {
  let firstOver12 = false;
  let secondOver12 = false;
  const numeric: string[] = [];
  for (const raw of texts) {
    const cleaned = String(raw ?? '').trim().replace(TIME_SUFFIX_RE, '').trim();
    const match = NUMERIC_RE.exec(cleaned);
    if (!match) continue;
    numeric.push(cleaned);
    if (Number(match[1]) > 12) firstOver12 = true;
    if (Number(match[2]) > 12) secondOver12 = true;
  }
  if (firstOver12 && !secondOver12) return 'dmy';
  if (secondOver12 && !firstOver12) return 'mdy';
  if (numeric.length < 3) return 'dmy';

  const inversions = (order: NumericDateOrder): number => {
    const ordinals = numeric
      .map((t) => parseStatementDate(t, order))
      .filter((p): p is DateParts => p !== null)
      .map(dateOrdinal);
    let up = 0;
    let down = 0;
    for (let i = 1; i < ordinals.length; i++) {
      if (ordinals[i] > ordinals[i - 1]) up++;
      else if (ordinals[i] < ordinals[i - 1]) down++;
    }
    // A statement runs one way throughout, so the smaller count is the disorder.
    return Math.min(up, down);
  };
  return inversions('mdy') < inversions('dmy') ? 'mdy' : 'dmy';
}

export interface ParsedAmount {
  /** Magnitude in paise, never negative. */
  paise: number;
  /** True when the text carried a minus sign or brackets. */
  negative: boolean;
  /** A Dr/Cr marker written in the same cell, when there was one. */
  marker?: 'dr' | 'cr';
}

const BLANK_AMOUNT_RE = /^(?:-+|–|—|nil|n\/a|na|\.+)?$/i;
const CURRENCY_RE = /₹|rs\.?|inr|\$|£|€/gi;
// Written with or without a space before it: "1,234.00 Cr", "1234.00Cr", "500.00 (Dr)".
const TRAILING_MARKER_RE = /\s*\(?(credit|debit|cr|dr)\.?\)?\s*$/i;
const LEADING_MARKER_RE = /^\s*\(?(credit|debit|cr|dr)\.?\)?\s*/i;
const DIGITS_RE = /^(\d*)(?:\.(\d+))?$/;

/**
 * Reads one amount. Returns null for a blank cell or text that is not a number,
 * so a caller can tell "nothing here" from zero.
 */
export function parseStatementAmount(text: string): ParsedAmount | null {
  let s = String(text ?? '').trim();
  if (BLANK_AMOUNT_RE.test(s)) return null;

  let marker: ParsedAmount['marker'];
  const markerMatch = TRAILING_MARKER_RE.exec(s) ?? LEADING_MARKER_RE.exec(s);
  if (markerMatch) {
    marker = markerMatch[1].toLowerCase().startsWith('c') ? 'cr' : 'dr';
    s = (s.slice(0, markerMatch.index) + s.slice(markerMatch.index + markerMatch[0].length)).trim();
  }

  s = s.replace(CURRENCY_RE, '').trim();

  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1).trim();
  }
  if (/^[-−–]/.test(s) || /[-−–]$/.test(s)) {
    negative = true;
    s = s.replace(/^[-−–]\s*/, '').replace(/\s*[-−–]$/, '');
  }
  s = s.replace(/^\+\s*/, '').replace(/[,\s]/g, '');

  const digits = DIGITS_RE.exec(s);
  if (!digits || (!digits[1] && !digits[2])) return null;

  const whole = digits[1] ? Number(digits[1]) : 0;
  const fraction = (digits[2] ?? '').padEnd(3, '0');
  let paise = whole * 100 + Number(fraction.slice(0, 2));
  if (Number(fraction[2]) >= 5) paise += 1;
  if (!Number.isSafeInteger(paise)) return null;

  return { paise, negative, ...(marker ? { marker } : {}) };
}

export function paiseToAmount(paise: number): number {
  return paise / 100;
}

/** Fixed two-decimal text for messages, e.g. 123456 -> "1234.56". */
export function formatPaise(paise: number): string {
  const sign = paise < 0 ? '-' : '';
  const abs = Math.abs(paise);
  return `${sign}${Math.floor(abs / 100)}.${String(abs % 100).padStart(2, '0')}`;
}

export type TokenKind = 'date' | 'money' | 'integer' | 'text';

const MONEY_SHAPE_RE = /\d\.\d{1,2}\b|\d,\d/;
const INTEGER_RE = /^\d+$/;

/**
 * Classifies one PDF text fragment so it can be matched to a column of the
 * right type. A bare integer is kept apart from money because it is as likely
 * to be a cheque or reference number as an amount.
 */
export function tokenKind(text: string): TokenKind {
  const s = String(text ?? '').trim();
  if (!s) return 'text';
  if (looksLikeDate(s)) return 'date';
  if (INTEGER_RE.test(s)) return 'integer';
  if (MONEY_SHAPE_RE.test(s) && parseStatementAmount(s) !== null) return 'money';
  return 'text';
}
