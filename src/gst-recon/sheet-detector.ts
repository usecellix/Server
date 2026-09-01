/**
 * Header/name scoring for PR vs GSTR sheet discovery (Server-side twin of client discovery).
 * Used for unit tests and future server-side conversational recon wiring.
 */

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

function norm(h: unknown): string {
  return String(h ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-./]+/g, ' ');
}

export function scoreSheetRole(
  sheetName: string,
  headers: string[],
  role: GstSheetRole,
): number {
  const name = norm(sheetName);
  const joined = headers.map(norm).join(' | ');
  let score = 0;
  const has = (...keys: string[]) => keys.some((k) => joined.includes(k) || name.includes(k));

  switch (role) {
    case 'GSTR2B':
      if (name.includes('2b') || name.includes('gstr2b')) score += 5;
      if (has('gstin of supplier', 'itc available', 'document type')) score += 3;
      if (has('taxable value') && has('gstin')) score += 2;
      break;
    case 'GSTR2A':
      if (name.includes('2a') || name.includes('gstr2a')) score += 5;
      break;
    case 'PURCHASE_REGISTER':
      if (name.includes('purchase')) score += 4;
      if (has('gstin') && has('invoice') && (has('taxable') || has('cgst'))) score += 3;
      if (name.includes('gstr') || name.includes('2b')) score -= 6;
      break;
    case 'SALES_REGISTER':
      if (name.includes('sales')) score += 4;
      break;
    case 'IMS':
      if (name.includes('ims')) score += 5;
      if (has('ims action', 'ims status')) score += 3;
      break;
    case 'GSTR1':
      if (name.includes('gstr1') || name.includes('gstr 1')) score += 5;
      break;
    default:
      break;
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
  missing: Array<'PURCHASE_REGISTER' | 'GSTR2B' | 'GSTR2A'>,
): string {
  const lines: string[] = [];
  if (missing.includes('GSTR2B')) {
    lines.push(
      "I couldn't find a GSTR-2B sheet in this workbook. Download GSTR-2B from the GST portal, add it as a sheet, then ask me again.",
    );
  }
  if (missing.includes('PURCHASE_REGISTER')) {
    lines.push(
      "I couldn't find a Purchase Register sheet with GSTIN/invoice/taxable columns in this workbook.",
    );
  }
  if (missing.includes('GSTR2A')) {
    lines.push(
      "I couldn't find a GSTR-2A sheet. Add your GSTR-2A download, then ask again.",
    );
  }
  return lines.join(' ');
}
