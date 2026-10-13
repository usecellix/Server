import { Channel } from '../../bank-recon/types';

export interface ParsedNarration {
  channel?: Channel;
  /** Cheque / instrument number (leading zeros stripped). */
  instrumentNo?: string;
  /** UTR / RTGS / NEFT / IMPS / UPI reference. */
  referenceNo?: string;
  counterparty?: string;
  invoiceNumbers: string[];
}

const NOISE_WORDS = new Set([
  'NEFT', 'RTGS', 'IMPS', 'UPI', 'CHQ', 'CHEQUE', 'CLG', 'ECS', 'NACH', 'ATM', 'POS', 'CASH',
  'DEP', 'DEPOSIT', 'PAYMENT', 'TRANSFER', 'TRF', 'TO', 'FROM', 'BY', 'INV', 'INVOICE', 'DR',
  'CR', 'PAID', 'RECEIVED', 'REF', 'NO', 'BANK', 'PAYTM', 'PYTM', 'OKAXIS', 'OKHDFCBANK', 'OKSBI',
  'YBL', 'IBL', 'AXL', 'UPIINT', 'MB', 'IB', 'NET', 'BANKING', 'INB',
]);

export function stripLeadingZeros(v: string): string {
  const t = v.replace(/^0+/, '');
  return t || v;
}

function detectChannel(n: string): Channel | undefined {
  if (/\bRTGS\b/.test(n)) return 'RTGS';
  if (/\bNEFT\b/.test(n)) return 'NEFT';
  if (/\bIMPS\b/.test(n)) return 'IMPS';
  if (/\bUPI\b/.test(n)) return 'UPI';
  if (/\b(NACH|ACH)\b/.test(n)) return 'NACH';
  if (/\bECS\b/.test(n)) return 'ECS';
  if (/\b(ATW|ATM|NWD|EAW)\b/.test(n)) return 'ATM';
  if (/\b(POS|ECOM|PURCHASE)\b/.test(n)) return 'POS';
  if (/\b(CHQ|CHEQUE|CLG|CLEARING|INST)\b/.test(n)) return 'CHQ';
  if (/\bCASH\b/.test(n)) return 'CASH';
  return undefined;
}

function pickCounterparty(segments: string[]): string | undefined {
  for (const seg of segments) {
    const t = seg.trim();
    if (t.length < 3) continue;
    if (/\d{4,}/.test(t)) continue; // refs, account numbers, invoice numbers
    if (!/[A-Za-z]{3}/.test(t)) continue;
    if (t.includes('@')) continue; // UPI handle
    const words = t.toUpperCase().split(/\s+/).filter(Boolean);
    if (words.every((w) => NOISE_WORDS.has(w))) continue;
    return t.replace(/\s+/g, ' ');
  }
  return undefined;
}

/**
 * Pull structured signals out of a messy bank/ledger narration:
 *   `NEFT-HDFCN52026-ABC TRADERS-INV 4521`, `UPI/412345678901/ABC/...`,
 *   `CHQ DEP 000123`, `ATW-1234XXXX`.
 * Deterministic, regex/rule based — never throws.
 */
export function parseNarration(narration: string): ParsedNarration {
  const text = String(narration ?? '').replace(/\s+/g, ' ').trim();
  const upper = text.toUpperCase();
  const out: ParsedNarration = { invoiceNumbers: [] };
  if (!text) return out;

  out.channel = detectChannel(upper);

  // UPI: 12-digit RRN
  const upi = upper.match(/\bUPI[-/ ]+(?:[A-Z]+[-/ ]+)?(\d{12})\b/) ?? upper.match(/\b(\d{12})\b/);
  if (upi && (out.channel === 'UPI' || out.channel === 'IMPS')) out.referenceNo = upi[1];

  // NEFT / RTGS / IMPS UTR: alphanumeric 10-22 chars containing a digit, after the channel.
  if (!out.referenceNo) {
    const utr = upper.match(/\b(?:NEFT|RTGS|IMPS)[-/ :]+(?:CR[-/ ]+|DR[-/ ]+)?([A-Z]{0,6}\d[A-Z0-9]{8,21})\b/);
    if (utr) out.referenceNo = utr[1];
  }
  if (!out.referenceNo) {
    const utr = upper.match(/\b(?:UTR|REF(?:ERENCE)?)\s*(?:NO\.?|#)?\s*[:\-]?\s*([A-Z0-9]{8,22})\b/);
    if (utr && /\d/.test(utr[1])) out.referenceNo = utr[1];
  }

  // Cheque number: 6 digits (occasionally 5-8) near a cheque keyword.
  const chq =
    upper.match(/\b(?:CHQ|CHEQUE|CHECK|CLG|INST(?:RUMENT)?)\.?\s*(?:DEP(?:OSIT)?\.?|NO\.?|NUM(?:BER)?|#)?\s*[:\-]?\s*(\d{3,8})\b/) ??
    upper.match(/\bCHQ\s*DEP\s*(\d{3,8})\b/);
  if (chq) {
    out.instrumentNo = stripLeadingZeros(chq[1]);
    out.channel = out.channel ?? 'CHQ';
  }

  // Embedded invoice numbers
  const invRe = /\b(?:INV(?:OICE)?\.?\s*(?:NO\.?|#)?\s*[:\-]?\s*)([A-Z0-9][A-Z0-9/\-]{2,19})\b/g;
  let im: RegExpExecArray | null;
  while ((im = invRe.exec(upper))) {
    if (/\d/.test(im[1])) out.invoiceNumbers.push(im[1]);
  }

  // Counterparty
  const segments = text.split(/[-/|]+|\s{2,}/);
  out.counterparty = pickCounterparty(segments);

  return out;
}

/** Normalize a reference token for cross-source comparison. */
export function normalizeRef(v: string | undefined): string {
  return stripLeadingZeros(String(v ?? '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
}

const STOP = new Set([
  'THE', 'AND', 'FOR', 'LTD', 'PVT', 'PRIVATE', 'LIMITED', 'M/S', 'MS', 'CO', 'COMPANY', 'LLP',
  'INC', 'CORP', 'TO', 'FROM', 'BY', 'NEFT', 'RTGS', 'IMPS', 'UPI', 'CHQ', 'CHEQUE', 'PAYMENT',
  'RECEIPT', 'RECEIVED', 'PAID', 'BEING', 'AGAINST', 'INV', 'INVOICE', 'BANK', 'TRANSFER', 'TRF',
  'DEP', 'DEPOSIT', 'NO', 'REF', 'A/C', 'ACCOUNT',
]);

export function narrationTokens(text: string): string[] {
  return String(text ?? '')
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, ' ')
    .split(/\s+/)
    .filter((t) => t.length >= 3 && !STOP.has(t) && !/^\d+$/.test(t));
}

/** Token-overlap similarity 0..1 (Dice on token sets, with prefix tolerance). */
export function narrationSimilarity(a: string, b: string): number {
  const ta = [...new Set(narrationTokens(a))];
  const tb = [...new Set(narrationTokens(b))];
  if (!ta.length || !tb.length) return 0;
  let hits = 0;
  const used = new Set<number>();
  for (const x of ta) {
    const idx = tb.findIndex(
      (y, i) => !used.has(i) && (x === y || (x.length >= 4 && y.length >= 4 && (x.startsWith(y) || y.startsWith(x)))),
    );
    if (idx >= 0) {
      used.add(idx);
      hits += 1;
    }
  }
  return (2 * hits) / (ta.length + tb.length);
}
