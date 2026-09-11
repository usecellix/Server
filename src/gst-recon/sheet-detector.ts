/**
 * Header-signature sheet discovery for Purchase/Sales GST reconciliation.
 * Spec §3 — aliases + distinguishing fields; no ML.
 */

import { RegisterSignature, SheetCandidate } from './types';

export type GstSheetRole =
  | 'PURCHASE_REGISTER'
  | 'SALES_REGISTER'
  | 'GSTR2B'
  | 'GSTR2A'
  | 'GSTR1'
  | 'IMS'
  | 'UNKNOWN';

export interface GstSheetScore {
  sheetName: string;
  role: GstSheetRole;
  score: number;
  headers: string[];
}

export interface HeaderSignatureRule {
  signature: RegisterSignature;
  requiredFields: string[];
  distinguishingFields: string[];
  minConfidence: number;
}

export const HEADER_ALIASES: Record<string, string[]> = {
  supplier_gstin: [
    'supplier gstin',
    'vendor gstin',
    'gstin of supplier',
    'gstin of the supplier',
  ],
  recipient_gstin: [
    'recipient gstin',
    'buyer gstin',
    'customer gstin',
    'gstin of recipient',
    'receiver gstin',
    'gstin of receiver',
    'gstin/uin of recipient',
  ],
  client_gstin: [
    'client gstin',
    'our gstin',
    'company gstin',
    'entity gstin',
    'gstin of registered person',
  ],
  invoice_number: [
    'invoice no',
    'bill no',
    'document no',
    'invoice number',
    'inv no',
    'voucher no',
    'document number',
  ],
  invoice_date: [
    'invoice date',
    'bill date',
    'document date',
    'voucher date',
    'inv date',
  ],
  taxable_value: [
    'taxable amount',
    'taxable value',
    'assessable value',
    'taxable amt',
    'net amount',
  ],
  igst: ['igst', 'igst amount'],
  cgst: ['cgst', 'cgst amount'],
  sgst: ['sgst', 'utgst', 'sgst amount'],
  cess: ['cess'],
  document_type: ['type', 'document type', 'invoice/note type', 'doc type', 'voucher type'],
  place_of_supply: ['pos', 'place of supply', 'state code'],
  irn: ['irn', 'e-invoice irn', 'invoice reference number'],
  eway_bill: ['e-way bill', 'eway bill no', 'ewb no'],
  supply_category: ['supply type', 'b2b/b2c', 'supply category', 'b2b', 'b2c'],
  itc_available: ['itc available', 'itc eligibility'],
  ims_action: ['ims action', 'ims status', 'recipient action'],
};

export const SIGNATURES: HeaderSignatureRule[] = [
  {
    signature: 'purchase_register',
    requiredFields: ['invoice_number', 'invoice_date', 'taxable_value'],
    distinguishingFields: ['supplier_gstin'],
    minConfidence: 0.7,
  },
  {
    signature: 'sales_register',
    requiredFields: ['invoice_number', 'invoice_date', 'taxable_value'],
    distinguishingFields: ['recipient_gstin', 'supply_category'],
    minConfidence: 0.7,
  },
  {
    signature: 'gstr_2b',
    requiredFields: ['supplier_gstin', 'invoice_number', 'invoice_date'],
    distinguishingFields: ['document_type', 'itc_available'],
    minConfidence: 0.75,
  },
  {
    signature: 'gstr_2a',
    requiredFields: ['supplier_gstin', 'invoice_number', 'invoice_date'],
    distinguishingFields: ['document_type'],
    minConfidence: 0.75,
  },
  {
    signature: 'gstr_1',
    requiredFields: ['invoice_number', 'invoice_date'],
    distinguishingFields: ['recipient_gstin', 'supply_category', 'irn'],
    minConfidence: 0.75,
  },
];

function normHeader(h: unknown): string {
  return String(h ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-./]+/g, ' ');
}

function fieldPresent(normalizedHeaders: string[], field: string): boolean {
  const aliases = (HEADER_ALIASES[field] ?? [field.replace(/_/g, ' ')]).map(normHeader);
  return aliases.some((alias) => {
    if (!alias) return false;
    // Short aliases (e.g. "type") require exact header match to avoid "supply type" false positives
    if (alias.length <= 4) {
      return normalizedHeaders.some((h) => h === alias);
    }
    return normalizedHeaders.some((h) => h === alias || h.includes(alias));
  });
}

function resolveFieldMap(
  normalizedHeaders: string[],
  originalHeaders: string[],
  fields: string[],
): Record<string, string> {
  const map: Record<string, string> = {};
  for (const field of fields) {
    const aliases = HEADER_ALIASES[field] ?? [field.replace(/_/g, ' ')];
    for (const alias of aliases) {
      const idx = normalizedHeaders.findIndex((h) => h === alias || h.includes(alias));
      if (idx >= 0) {
        map[field] = originalHeaders[idx] ?? alias;
        break;
      }
    }
  }
  return map;
}

export function detectSheetSignature(headers: string[]): {
  signature: RegisterSignature | null;
  confidence: number;
  fieldMap: Record<string, string>;
} {
  const norms = headers.map(normHeader);
  let best: {
    signature: RegisterSignature;
    confidence: number;
    fieldMap: Record<string, string>;
  } | null = null;

  for (const rule of SIGNATURES) {
    const requiredHits = rule.requiredFields.filter((f) => fieldPresent(norms, f)).length;
    if (requiredHits < rule.requiredFields.length) continue;

    const distHits = rule.distinguishingFields.filter((f) => fieldPresent(norms, f)).length;
    const denom = rule.requiredFields.length + rule.distinguishingFields.length;
    let confidence = (requiredHits + distHits) / Math.max(denom, 1);

    // Lookalike penalties / boosts
    if (rule.signature === 'purchase_register' && fieldPresent(norms, 'recipient_gstin')) {
      confidence -= 0.35;
    }
    if (rule.signature === 'sales_register' && fieldPresent(norms, 'supplier_gstin')) {
      confidence -= 0.1;
    }
    if (rule.signature === 'sales_register' && fieldPresent(norms, 'irn')) {
      confidence -= 0.25;
    }
    if (rule.signature === 'gstr_2b' && fieldPresent(norms, 'itc_available')) {
      confidence = Math.min(1, confidence + 0.05);
    }
    if (
      rule.signature === 'gstr_2a' &&
      fieldPresent(norms, 'recipient_gstin') &&
      !fieldPresent(norms, 'itc_available')
    ) {
      confidence -= 0.2;
    }
    if (
      rule.signature === 'gstr_1' &&
      fieldPresent(norms, 'irn')
    ) {
      confidence = Math.min(1, confidence + 0.2);
    }
    if (
      rule.signature === 'gstr_1' &&
      fieldPresent(norms, 'supplier_gstin') &&
      !fieldPresent(norms, 'irn')
    ) {
      confidence -= 0.3;
    }

    if (confidence < rule.minConfidence) continue;

    const fieldMap = resolveFieldMap(norms, headers, [
      ...rule.requiredFields,
      ...rule.distinguishingFields,
    ]);

    if (!best || confidence > best.confidence) {
      best = { signature: rule.signature, confidence, fieldMap };
    }
  }

  if (!best) {
    return { signature: null, confidence: 0, fieldMap: {} };
  }
  return best;
}

export interface ResolveResult {
  found: SheetCandidate[];
  status: 'resolved' | 'ambiguous' | 'not_found';
}

export function resolveSheetsForRecon(
  allSheets: { sheetName: string; headers: string[] }[],
  needed: RegisterSignature[],
): Record<RegisterSignature, ResolveResult> {
  const out = {} as Record<RegisterSignature, ResolveResult>;
  for (const sig of needed) {
    const matches: SheetCandidate[] = [];
    for (const sheet of allSheets) {
      const detected = detectSheetSignature(sheet.headers);
      const nameBoost = nameHintBoost(sheet.sheetName, sig);
      let confidence = detected.signature === sig ? detected.confidence : 0;
      // Also allow name-driven candidates when headers partially match
      if (detected.signature !== sig) {
        const rule = SIGNATURES.find((r) => r.signature === sig);
        if (rule) {
          const norms = sheet.headers.map(normHeader);
          const requiredHits = rule.requiredFields.filter((f) =>
            fieldPresent(norms, f),
          ).length;
          if (requiredHits === rule.requiredFields.length && nameBoost >= 0.2) {
            confidence = Math.max(confidence, 0.7 + nameBoost);
          }
        }
      } else {
        confidence = Math.min(1, confidence + nameBoost);
      }
      const minConf =
        SIGNATURES.find((r) => r.signature === sig)?.minConfidence ?? 0.7;
      if (confidence >= minConf) {
        matches.push({
          sheetName: sheet.sheetName,
          headers: sheet.headers,
          headerRowIndex: 1,
          confidence,
        });
      }
    }
    matches.sort((a, b) => b.confidence - a.confidence);
    if (matches.length === 0) {
      out[sig] = { found: [], status: 'not_found' };
    } else if (matches.length === 1) {
      out[sig] = { found: matches, status: 'resolved' };
    } else {
      const top = matches[0].confidence;
      const ties = matches.filter((m) => Math.abs(m.confidence - top) < 0.05);
      if (ties.length > 1) {
        out[sig] = { found: ties, status: 'ambiguous' };
      } else {
        out[sig] = { found: [matches[0]], status: 'resolved' };
      }
    }
  }
  return out;
}

function nameHintBoost(sheetName: string, sig: RegisterSignature): number {
  const name = normHeader(sheetName);
  switch (sig) {
    case 'purchase_register':
      if (name.includes('gstr') || name.includes('2b') || name.includes('2a')) return -0.4;
      if (name.includes('purchase') || /\bpr\b/.test(name)) return 0.2;
      return 0;
    case 'sales_register':
      if (name.includes('gstr')) return -0.3;
      if (name.includes('sales') || name.includes('outward')) return 0.2;
      return 0;
    case 'gstr_2b':
      if (name.includes('2b') || name.includes('gstr2b')) return 0.25;
      return 0;
    case 'gstr_2a':
      if (name.includes('2a') || name.includes('gstr2a')) return 0.25;
      return 0;
    case 'gstr_1':
      if (name.includes('gstr1') || name.includes('gstr 1') || /gstr-?1/.test(name))
        return 0.25;
      return 0;
    default:
      return 0;
  }
}

export const SHEET_MESSAGES: Record<RegisterSignature, { notFound: string }> = {
  purchase_register: {
    notFound:
      "I couldn't find a Purchase Register sheet in this workbook. Please check your Purchase Register export and add it as a sheet, then ask me again.",
  },
  sales_register: {
    notFound:
      "I couldn't find a Sales Register sheet in this workbook. Please check your Sales Register export and add it as a sheet, then ask me again.",
  },
  gstr_2b: {
    notFound:
      "I couldn't find a GSTR-2B sheet in this workbook. Please download GSTR-2B from the GST portal for this client's GSTIN and add it as a sheet, then ask me again.",
  },
  gstr_2a: {
    notFound:
      "I couldn't find a GSTR-2A sheet in this workbook. Please download GSTR-2A from the GST portal for this client's GSTIN and add it as a sheet, then ask me again.",
  },
  gstr_1: {
    notFound:
      "I couldn't find a GSTR-1 sheet in this workbook. Please download the GSTR-1 export from the GST portal for this client's GSTIN and add it as a sheet, then ask me again.",
  },
};

export function ambiguousSheetMessage(
  signature: RegisterSignature,
  candidates: SheetCandidate[],
): string {
  const label = {
    purchase_register: 'purchase register',
    sales_register: 'sales register',
    gstr_2b: 'GSTR-2B sheet',
    gstr_2a: 'GSTR-2A sheet',
    gstr_1: 'GSTR-1 sheet',
  }[signature];
  const names = candidates.map((c) => `'${c.sheetName}'`).join(' and ');
  return `I found more than one possible ${label}: ${names} — which one should I use?`;
}

/** Legacy scoring helpers kept for existing tests / IMS role. */
export function scoreSheetRole(
  sheetName: string,
  headers: string[],
  role: GstSheetRole,
): number {
  const sigMap: Partial<Record<GstSheetRole, RegisterSignature>> = {
    PURCHASE_REGISTER: 'purchase_register',
    SALES_REGISTER: 'sales_register',
    GSTR2B: 'gstr_2b',
    GSTR2A: 'gstr_2a',
    GSTR1: 'gstr_1',
  };
  const sig = sigMap[role];
  if (sig) {
    const detected = detectSheetSignature(headers);
    const boost = nameHintBoost(sheetName, sig);
    if (detected.signature === sig) {
      return Math.round((detected.confidence + boost) * 10);
    }
    // Partial: required fields only
    const rule = SIGNATURES.find((r) => r.signature === sig);
    if (rule) {
      const norms = headers.map(normHeader);
      const requiredHits = rule.requiredFields.filter((f) => fieldPresent(norms, f)).length;
      if (requiredHits === rule.requiredFields.length) {
        return Math.round((0.7 + boost) * 10);
      }
    }
  }

  // IMS / fallback name heuristics
  const name = normHeader(sheetName);
  const joined = headers.map(normHeader).join(' | ');
  let score = 0;
  const has = (...keys: string[]) => keys.some((k) => joined.includes(k) || name.includes(k));
  if (role === 'IMS') {
    if (name.includes('ims')) score += 5;
    if (has('ims action', 'ims status')) score += 3;
  }
  return score;
}

export function pickBestSheet(
  sheets: Array<{ sheetName: string; headers: string[] }>,
  role: GstSheetRole,
  minScore = 3,
): { best: GstSheetScore | null; ties: string[] } {
  const scored = sheets
    .map((s) => ({
      sheetName: s.sheetName,
      role,
      score: scoreSheetRole(s.sheetName, s.headers, role),
      headers: s.headers,
    }))
    .filter((s) => s.score >= minScore)
    .sort((a, b) => b.score - a.score);

  if (!scored.length) return { best: null, ties: [] };
  const top = scored[0].score;
  const ties = scored.filter((s) => s.score === top).map((s) => s.sheetName);
  if (ties.length > 1) return { best: null, ties };
  return { best: scored[0], ties: [] };
}

export function missingSheetChatMessage(
  missing: Array<'PURCHASE_REGISTER' | 'GSTR2B' | 'GSTR2A' | 'SALES_REGISTER' | 'GSTR1'>,
): string {
  const lines: string[] = [];
  for (const m of missing) {
    if (m === 'GSTR2B') lines.push(SHEET_MESSAGES.gstr_2b.notFound);
    if (m === 'GSTR2A') lines.push(SHEET_MESSAGES.gstr_2a.notFound);
    if (m === 'GSTR1') lines.push(SHEET_MESSAGES.gstr_1.notFound);
    if (m === 'PURCHASE_REGISTER') lines.push(SHEET_MESSAGES.purchase_register.notFound);
    if (m === 'SALES_REGISTER') lines.push(SHEET_MESSAGES.sales_register.notFound);
  }
  return lines.join('\n\n') || 'Required sheets were not found.';
}
