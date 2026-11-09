/** Small shared helpers for deterministic lanes. Pure, no I/O. */

export function columnLetter(index: number): string {
  let n = index;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

export function letterToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

const NOT_LETTERS = /^(of|in|is|to|or|for|the|has|are|an|as|on|by|at|it|be|if|we|my|named|called|with|from|that|which|and)$/i;

/** All "column X" / "columns X and Y" letters in the message, in order. */
export function columnsFromMessage(message: string): number[] {
  const out: number[] = [];
  const re = /\bcolumns?\s+([A-Za-z]{1,2}(?:\s*(?:,|and|&)\s*[A-Za-z]{1,2}\b)*)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(message))) {
    for (const part of m[1].split(/\s*(?:,|and|&)\s*/i)) {
      if (/^[A-Za-z]{1,2}$/.test(part) && !NOT_LETTERS.test(part)) out.push(letterToIndex(part));
    }
  }
  return out;
}

/** A cell address like "F2" / "D2" mentioned as a target ("in F2", "at D2"). */
export function targetCellFromMessage(message: string): { col: number; row: number } | null {
  const m = /\b(?:in|at|into|to)\s+(?:cell\s+)?\$?([A-Za-z]{1,2})\$?(\d{1,7})\b/i.exec(message);
  if (!m) return null;
  return { col: letterToIndex(m[1]), row: Number(m[2]) - 1 };
}

export const str = (v: unknown): string => (v === null || v === undefined ? '' : String(v));
export const isBlank = (v: unknown): boolean => v === null || v === undefined || v === '';

export function lastUsedColumn(values: unknown[][]): number {
  let last = -1;
  for (const row of values) {
    if (!Array.isArray(row)) continue;
    for (let c = row.length - 1; c > last; c -= 1) {
      if (str(row[c]).trim() !== '') {
        last = c;
        break;
      }
    }
  }
  return last;
}

export function lastUsedRow(values: unknown[][]): number {
  for (let r = values.length - 1; r >= 0; r -= 1) {
    if (Array.isArray(values[r]) && values[r].some((v) => str(v).trim() !== '')) return r;
  }
  return -1;
}

/** Index of the first header cell matching the pattern, or -1. */
export function findHeader(header: unknown[], pattern: RegExp): number {
  return header.findIndex((h) => pattern.test(str(h)));
}

/** Header text named in the message ("from Invoice Date", "by Party") → column index. */
export function headerNamedInMessage(header: unknown[], message: string): number {
  const lower = message.toLowerCase();
  let best = -1;
  let bestLen = 0;
  header.forEach((h, i) => {
    const name = str(h).trim().toLowerCase();
    if (name.length >= 3 && lower.includes(name) && name.length > bestLen) {
      best = i;
      bestLen = name.length;
    }
  });
  return best;
}

// ---- Dates (Excel serials and Indian text dates) ----

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

/** UTC-midnight ms for an Excel serial, ISO date, dd-mm-yyyy, dd/mm/yy, dd.mm.yy or "31 Mar 2026". */
export function parseDateValue(v: unknown): number | null {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return Date.UTC(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number' && Number.isFinite(v)) {
    return v > 20_000 && v < 80_000 ? Math.round((Math.floor(v) - 25_569) * 86_400_000) : null;
  }
  const t = str(v).trim();
  if (!t) return null;
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(t);
  if (m) {
    const year = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    if (+m[2] < 1 || +m[2] > 12 || +m[1] < 1 || +m[1] > 31) return null;
    return Date.UTC(year, +m[2] - 1, +m[1]);
  }
  m = /^(\d{1,2})[-\s]([A-Za-z]{3})[A-Za-z]*[-\s,]+(\d{2,4})$/.exec(t);
  if (m && MONTHS.includes(m[2].toLowerCase())) {
    const year = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    return Date.UTC(year, MONTHS.indexOf(m[2].toLowerCase()), +m[1]);
  }
  return null;
}

export const toExcelSerial = (ms: number): number => Math.round(ms / 86_400_000) + 25_569;
export const daysBetween = (fromMs: number, toMs: number): number => Math.round((toMs - fromMs) / 86_400_000);

export function formatDdMmYyyy(ms: number): string {
  const d = new Date(ms);
  return `${String(d.getUTCDate()).padStart(2, '0')}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${d.getUTCFullYear()}`;
}

/** "as on 31-03-2026", "as of 31/03/2026", "and 31-03-2026", "on 31 Mar 2026" → UTC ms. */
export function dateInMessage(message: string): number | null {
  const m =
    /\b(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[\s-][A-Za-z]{3,9}[\s-,]+\d{2,4})\b/.exec(message);
  return m ? parseDateValue(m[1]) : null;
}

export function excelDateLiteral(ms: number): string {
  const d = new Date(ms);
  return `DATE(${d.getUTCFullYear()},${d.getUTCMonth() + 1},${d.getUTCDate()})`;
}

// ---- Indian number display ----

export function indianNumber(n: number, decimals = 0): string {
  const neg = n < 0;
  const [int, frac] = Math.abs(n).toFixed(decimals).split('.');
  const last3 = int.slice(-3);
  const rest = int.slice(0, -3);
  const grouped = rest ? `${rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',')},${last3}` : last3;
  return `${neg ? '-' : ''}${grouped}${frac ? `.${frac}` : ''}`;
}

export const toNumber = (v: unknown): number | null => {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  const t = str(v).replace(/[,\s₹]/g, '');
  if (!t || !/^[-+]?\d+(\.\d+)?$/.test(t)) return null;
  return Number(t);
};
