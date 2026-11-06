/**
 * P1.4 — "Validate the GSTINs in column B" (A9.2), deterministic lane.
 *
 * No LLM: reads the active sheet, validates every GSTIN with the shared B10
 * module and proposes two new columns (GSTIN Status, GSTIN State) as an
 * ordinary BATCH_SET — so it goes through preview → Accept → undo like every
 * other write. The original GSTIN column is never touched and no identifier is
 * ever auto-corrected.
 */
import { BatchSetOperation, SheetActionPayload } from '../types/sheet-actions.types';
import {
  GSTIN_STATUS_LABEL,
  GstinStatus,
  normalizeIdentifier,
  validateGstin,
} from './india-validators.util';

export const GSTIN_STATUS_HEADER = 'GSTIN Status';
export const GSTIN_STATE_HEADER = 'GSTIN State';

const VERB = /\b(validate|validat(?:e|ing)|verify|check|audit|test)\b/i;
const GSTIN_WORD = /\bgst-?ins?\b/i;
/** Messages that are about reconciliation / matching, not column validation. */
const OTHER_JOB = /\b(recon\w*|match\w*|gstr[-\s]?\w*|2a|2b|compare|vs\.?|versus|highlight|delete|remove|filter|sort)\b/i;

export function detectValidateGstinIntent(message: string): boolean {
  if (!message) return false;
  return VERB.test(message) && GSTIN_WORD.test(message) && !OTHER_JOB.test(message);
}

export function columnLetter(index: number): string {
  let n = index;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

function letterToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function explicitColumnFromMessage(message: string): number | null {
  const m = /\bcolumn\s+([A-Za-z]{1,3})\b/i.exec(message);
  if (!m) return null;
  // "column named …" / "column of …" are words, not letters.
  if (/^(named|of|with|for|that|which|called|and|in|to|from)$/i.test(m[1])) return null;
  return letterToIndex(m[1]);
}

const looksLikeGstin = (v: unknown): boolean => {
  const s = normalizeIdentifier(v);
  return s.length === 15 && /^[0-9]{2}/.test(s);
};

function lastUsedColumn(sheetData: unknown[][]): number {
  let last = -1;
  for (const row of sheetData) {
    if (!Array.isArray(row)) continue;
    for (let c = row.length - 1; c > last; c -= 1) {
      if (String(row[c] ?? '').trim() !== '') {
        last = c;
        break;
      }
    }
  }
  return last;
}

/** Pick the GSTIN column: explicit letter → header containing "gstin" → best content match. */
export function findGstinColumn(
  sheetData: unknown[][],
  headerRowIndex: number,
  message: string,
): number | null {
  const explicit = explicitColumnFromMessage(message);
  if (explicit !== null && explicit >= 0) return explicit;

  const header = sheetData[headerRowIndex] ?? [];
  const byHeader = header.findIndex((h) => GSTIN_WORD.test(String(h ?? '')));
  if (byHeader >= 0) return byHeader;

  const width = lastUsedColumn(sheetData) + 1;
  let best = -1;
  let bestRatio = 0;
  for (let c = 0; c < width; c += 1) {
    let filled = 0;
    let hits = 0;
    for (let r = headerRowIndex + 1; r < sheetData.length; r += 1) {
      const v = sheetData[r]?.[c];
      if (String(v ?? '').trim() === '') continue;
      filled += 1;
      if (looksLikeGstin(v)) hits += 1;
    }
    if (filled > 0 && hits / filled > bestRatio) {
      bestRatio = hits / filled;
      best = c;
    }
  }
  return bestRatio >= 0.5 ? best : null;
}

export interface GstinValidationCounts {
  checked: number;
  valid: number;
  invalid_format: number;
  invalid_state: number;
  invalid_pan: number;
  checksum_fail: number;
}

export interface GstinLaneResult {
  kind: 'actions' | 'answer';
  answer: string;
  explanation?: string;
  actions?: SheetActionPayload[];
  counts?: GstinValidationCounts;
}

export function buildGstinValidationPlan(
  sheetData: unknown[][],
  headerRowIndex: number,
  sheetName: string,
  message: string,
): GstinLaneResult {
  const col = findGstinColumn(sheetData, headerRowIndex, message);
  if (col === null) {
    return {
      kind: 'answer',
      answer:
        "I couldn't find a GSTIN column. Tell me which column to check (for example: *Validate the GSTINs in column B*).",
    };
  }

  const header = sheetData[headerRowIndex] ?? [];
  const existingStatus = header.findIndex((h) => String(h ?? '').trim() === GSTIN_STATUS_HEADER);
  const existingState = header.findIndex((h) => String(h ?? '').trim() === GSTIN_STATE_HEADER);
  let nextFree = lastUsedColumn(sheetData) + 1;
  const statusCol = existingStatus >= 0 ? existingStatus : nextFree++;
  const stateCol = existingState >= 0 ? existingState : nextFree++;

  const headerRowNumber = headerRowIndex + 1;
  const operations: BatchSetOperation[] = [
    { address: `${columnLetter(statusCol)}${headerRowNumber}`, value: GSTIN_STATUS_HEADER },
    { address: `${columnLetter(stateCol)}${headerRowNumber}`, value: GSTIN_STATE_HEADER },
  ];

  const counts: GstinValidationCounts = {
    checked: 0,
    valid: 0,
    invalid_format: 0,
    invalid_state: 0,
    invalid_pan: 0,
    checksum_fail: 0,
  };
  const review: string[] = [];

  for (let r = headerRowIndex + 1; r < sheetData.length; r += 1) {
    const raw = sheetData[r]?.[col];
    if (String(raw ?? '').trim() === '') continue; // blank → left blank, not counted
    const result = validateGstin(raw);
    counts.checked += 1;
    counts[result.status] += 1;
    const rowNumber = r + 1;
    operations.push({
      address: `${columnLetter(statusCol)}${rowNumber}`,
      value: GSTIN_STATUS_LABEL[result.status],
    });
    operations.push({
      address: `${columnLetter(stateCol)}${rowNumber}`,
      value: result.stateName ?? '',
    });
    if (result.status !== 'valid' && review.length < 10) {
      review.push(`row ${rowNumber} — ${String(raw).trim()} → ${GSTIN_STATUS_LABEL[result.status as GstinStatus]}`);
    }
  }

  if (counts.checked === 0) {
    return {
      kind: 'answer',
      answer: `Column ${columnLetter(col)} has no GSTIN values below the header row, so there is nothing to validate.`,
    };
  }

  const parts = [
    `**${counts.valid} valid**`,
    counts.checksum_fail ? `${counts.checksum_fail} checksum fail` : '',
    counts.invalid_format ? `${counts.invalid_format} invalid format` : '',
    counts.invalid_state ? `${counts.invalid_state} invalid state` : '',
    counts.invalid_pan ? `${counts.invalid_pan} invalid PAN segment` : '',
  ].filter(Boolean);

  const lines = [
    `Checked **${counts.checked}** GSTINs in column ${columnLetter(col)}: ${parts.join(', ')}.`,
    `I'll add **${GSTIN_STATUS_HEADER}** (column ${columnLetter(statusCol)}) and **${GSTIN_STATE_HEADER}** (column ${columnLetter(stateCol)}). Your original GSTIN column is not changed, and nothing is auto-corrected.`,
  ];
  if (review.length) {
    lines.push('', 'Flagged for review:', ...review.map((l) => `- ${l}`));
  }
  lines.push('', '_Status reflects the GSTIN format, state code, PAN segment and check digit — not whether the registration is active on the GST portal._');

  return {
    kind: 'actions',
    answer: lines.join('\n'),
    explanation: `Validate ${counts.checked} GSTINs and add status and state columns.`,
    actions: [{ type: 'BATCH_SET', sheetName, operations }] as SheetActionPayload[],
    counts,
  };
}
