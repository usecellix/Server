/**
 * B10 — Indian identifier validators (one module; replaces the duplicated
 * GSTIN regexes in gst-recon/gstin-validator.ts and domain-tools/gst/normalize-invoice.ts).
 *
 * Deterministic, no LLM, no I/O. NEVER auto-corrects an identifier — it only
 * classifies it. Statutory lists (state codes) are kept as data here until the
 * rule-table registry (B5) exists; see TODO(B5).
 */

/* ------------------------------------------------------------------ */
/* State codes                                                         */
/* ------------------------------------------------------------------ */

/** TODO(B5): move to the versioned rule-table registry once it exists. */
export const GST_STATE_CODES: Readonly<Record<string, string>> = {
  '01': 'Jammu & Kashmir',
  '02': 'Himachal Pradesh',
  '03': 'Punjab',
  '04': 'Chandigarh',
  '05': 'Uttarakhand',
  '06': 'Haryana',
  '07': 'Delhi',
  '08': 'Rajasthan',
  '09': 'Uttar Pradesh',
  '10': 'Bihar',
  '11': 'Sikkim',
  '12': 'Arunachal Pradesh',
  '13': 'Nagaland',
  '14': 'Manipur',
  '15': 'Mizoram',
  '16': 'Tripura',
  '17': 'Meghalaya',
  '18': 'Assam',
  '19': 'West Bengal',
  '20': 'Jharkhand',
  '21': 'Odisha',
  '22': 'Chhattisgarh',
  '23': 'Madhya Pradesh',
  '24': 'Gujarat',
  '25': 'Daman & Diu (merged into 26)',
  '26': 'Dadra & Nagar Haveli and Daman & Diu',
  '27': 'Maharashtra',
  '28': 'Andhra Pradesh (before 2014 split)',
  '29': 'Karnataka',
  '30': 'Goa',
  '31': 'Lakshadweep',
  '32': 'Kerala',
  '33': 'Tamil Nadu',
  '34': 'Puducherry',
  '35': 'Andaman & Nicobar Islands',
  '36': 'Telangana',
  '37': 'Andhra Pradesh',
  '38': 'Ladakh',
  '97': 'Other Territory',
  '99': 'Centre Jurisdiction',
};

export function gstStateName(code: string): string | undefined {
  return GST_STATE_CODES[code];
}

/* ------------------------------------------------------------------ */
/* PAN                                                                 */
/* ------------------------------------------------------------------ */

export const PAN_REGEX = /^[A-Z]{5}[0-9]{4}[A-Z]$/;

/** 4th character of a PAN = holder type. */
export const PAN_ENTITY_TYPES: Readonly<Record<string, string>> = {
  P: 'Individual',
  C: 'Company',
  H: 'HUF',
  F: 'Firm / LLP',
  A: 'AOP',
  B: 'BOI',
  T: 'Trust',
  L: 'Local Authority',
  J: 'Artificial Juridical Person',
  G: 'Government',
};

export type PanStatus = 'valid' | 'invalid_format' | 'invalid_entity_type';

export interface PanValidation {
  status: PanStatus;
  pan: string;
  entityType?: string;
}

export function normalizeIdentifier(raw: unknown): string {
  return String(raw ?? '')
    .replace(/[\s ]+/g, '')
    .toUpperCase();
}

export function validatePan(raw: unknown): PanValidation {
  const pan = normalizeIdentifier(raw);
  if (!PAN_REGEX.test(pan)) return { status: 'invalid_format', pan };
  const entityType = PAN_ENTITY_TYPES[pan[3]];
  if (!entityType) return { status: 'invalid_entity_type', pan };
  return { status: 'valid', pan, entityType };
}

/** Find PAN-shaped tokens inside free text (e.g. "Vendor ABC PAN AABCA1234F paid"). */
export function extractPans(text: unknown): string[] {
  const found = String(text ?? '')
    .toUpperCase()
    .match(/\b[A-Z]{5}[0-9]{4}[A-Z]\b/g);
  return found ? Array.from(new Set(found)) : [];
}

/* ------------------------------------------------------------------ */
/* GSTIN                                                               */
/* ------------------------------------------------------------------ */

/** Format only: 2 digits state, PAN (10), entity digit, 'Z', check char. */
export const GSTIN_FORMAT_REGEX = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;

const GSTIN_CHARSET = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';

/** Check character for the first 14 chars of a GSTIN (mod-36 weighted sum). */
export function gstinCheckChar(first14: string): string | null {
  if (first14.length !== 14) return null;
  let sum = 0;
  for (let i = 0; i < 14; i += 1) {
    const value = GSTIN_CHARSET.indexOf(first14[i]);
    if (value < 0) return null;
    const product = value * (i % 2 === 0 ? 1 : 2);
    sum += Math.floor(product / 36) + (product % 36);
  }
  return GSTIN_CHARSET[(36 - (sum % 36)) % 36];
}

export type GstinStatus =
  | 'valid'
  | 'invalid_format'
  | 'invalid_state'
  | 'invalid_pan'
  | 'checksum_fail';

export interface GstinValidation {
  status: GstinStatus;
  gstin: string;
  stateCode?: string;
  stateName?: string;
  pan?: string;
  panEntityType?: string;
  /** Present when status === 'checksum_fail': what the check char should be. */
  expectedCheckChar?: string;
}

export const GSTIN_STATUS_LABEL: Readonly<Record<GstinStatus, string>> = {
  valid: 'Valid',
  invalid_format: 'Invalid format',
  invalid_state: 'Invalid state',
  invalid_pan: 'Invalid PAN segment',
  checksum_fail: 'Checksum fail',
};

/**
 * Order of checks: format → state code → PAN segment → checksum. The first
 * failing check names the status. Never mutates or "fixes" the input.
 */
export function validateGstin(raw: unknown): GstinValidation {
  const gstin = normalizeIdentifier(raw);
  if (!GSTIN_FORMAT_REGEX.test(gstin)) return { status: 'invalid_format', gstin };

  const stateCode = gstin.slice(0, 2);
  const stateName = GST_STATE_CODES[stateCode];
  if (!stateName) return { status: 'invalid_state', gstin, stateCode };

  const pan = gstin.slice(2, 12);
  const panResult = validatePan(pan);
  if (panResult.status !== 'valid') {
    return { status: 'invalid_pan', gstin, stateCode, stateName, pan };
  }

  const expected = gstinCheckChar(gstin.slice(0, 14));
  if (expected !== gstin[14]) {
    return {
      status: 'checksum_fail',
      gstin,
      stateCode,
      stateName,
      pan,
      panEntityType: panResult.entityType,
      expectedCheckChar: expected ?? undefined,
    };
  }

  return { status: 'valid', gstin, stateCode, stateName, pan, panEntityType: panResult.entityType };
}

/** Find GSTIN-shaped tokens inside free text. Returned as found (not validated). */
export function extractGstins(text: unknown): string[] {
  const found = String(text ?? '')
    .toUpperCase()
    .match(/\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b/g);
  return found ? Array.from(new Set(found)) : [];
}

/* ------------------------------------------------------------------ */
/* TAN / IFSC / Udyam / HSN-SAC                                        */
/* ------------------------------------------------------------------ */

export const TAN_REGEX = /^[A-Z]{4}[0-9]{5}[A-Z]$/;
export function validateTan(raw: unknown): boolean {
  return TAN_REGEX.test(normalizeIdentifier(raw));
}

/** 4 letters (bank), '0', then 6 alphanumerics (branch). */
export const IFSC_REGEX = /^[A-Z]{4}0[A-Z0-9]{6}$/;
export function validateIfsc(raw: unknown): boolean {
  return IFSC_REGEX.test(normalizeIdentifier(raw));
}

/** UDYAM-XX-00-0000000 */
export const UDYAM_REGEX = /^UDYAM-[A-Z]{2}-[0-9]{2}-[0-9]{7}$/;
export function validateUdyam(raw: unknown): boolean {
  return UDYAM_REGEX.test(String(raw ?? '').trim().toUpperCase());
}

export type HsnSacKind = 'HSN' | 'SAC';
export interface HsnSacValidation {
  valid: boolean;
  kind?: HsnSacKind;
  code: string;
}

/**
 * HSN: 2, 4, 6 or 8 digits (GSTR-1 Table 12 requires 4/6/8 by turnover).
 * SAC: 6 digits starting with 99.
 */
export function validateHsnSac(raw: unknown): HsnSacValidation {
  const code = String(raw ?? '').replace(/\s+/g, '');
  if (!/^[0-9]+$/.test(code)) return { valid: false, code };
  if (code.length === 6 && code.startsWith('99')) return { valid: true, kind: 'SAC', code };
  if ([2, 4, 6, 8].includes(code.length)) return { valid: true, kind: 'HSN', code };
  return { valid: false, code };
}

/* ------------------------------------------------------------------ */
/* Invoice-number normaliser (shared by fuzzy matching, dedup, recon)  */
/* ------------------------------------------------------------------ */

const INVOICE_PREFIX_WORDS = new Set([
  'INV', 'INVOICE', 'BILL', 'NO', 'NUM', 'NUMBER', 'TAX', 'REF', 'REFNO', 'VCH', 'VOUCHER',
]);

function isFyPair(a: string, b: string): boolean {
  if (!/^[0-9]+$/.test(a) || !/^[0-9]+$/.test(b)) return false;
  const na = Number(a);
  const nb = Number(b);
  if (a.length === 2 && b.length === 2) return (na + 1) % 100 === nb;
  if (a.length === 4 && b.length === 2) return a.startsWith('20') && ((na % 100) + 1) % 100 === nb;
  if (a.length === 4 && b.length === 4) return a.startsWith('20') && na + 1 === nb;
  return false;
}

/**
 * Reduce an invoice reference (or a narration containing one) to a comparable
 * key: strips prefixes (INV / Bill No / #), financial-year parts (24-25),
 * trailing "dt <date>", punctuation and leading zeros.
 *
 *   "INV/24-25/0233"          → "233"
 *   "Inv#233"                 → "233"
 *   "0233"                    → "233"
 *   "bill no 233 dt 12.04.24" → "233"
 *   "OS-900" / "OS900"        → "OS900"
 *
 * The raw value is never altered; callers keep it next to the key.
 */
export function normalizeInvoiceKey(raw: unknown): string {
  let s = String(raw ?? '').toUpperCase().trim();
  if (!s) return '';

  // Cut a trailing date clause: "... DT 12.04.24", "... DATED 12/04/2024".
  const dateCut = s.search(/\b(?:DT|DATED|DATE)\b/);
  if (dateCut > 0) s = s.slice(0, dateCut);

  let tokens = s.split(/[^A-Z0-9]+/).filter(Boolean);

  // Drop leading prefix words (INV, BILL, NO, ...).
  while (tokens.length > 1 && INVOICE_PREFIX_WORDS.has(tokens[0])) tokens.shift();
  // Drop prefix words glued to nothing else but also in the middle ("BILL NO 233").
  tokens = tokens.filter((t, i) => !(i < tokens.length - 1 && INVOICE_PREFIX_WORDS.has(t)));

  // Drop financial-year pairs such as 24-25 / 2024-25.
  const withoutFy: string[] = [];
  for (let i = 0; i < tokens.length; i += 1) {
    if (i + 1 < tokens.length && isFyPair(tokens[i], tokens[i + 1])) {
      i += 1;
      continue;
    }
    withoutFy.push(tokens[i]);
  }
  if (withoutFy.length > 0) tokens = withoutFy;

  // Leading zeros off purely numeric tokens ("0233" → "233", "000" → "0").
  tokens = tokens.map((t) => (/^[0-9]+$/.test(t) ? t.replace(/^0+(?=[0-9])/, '') : t));

  return tokens.join('');
}
