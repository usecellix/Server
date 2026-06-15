import { SheetActionPayload } from '../../excel-ai/types/sheet-actions.types';
import { SubTask, WorkbookContext } from '../types/agent.types';
import { normalizeSortRangeAddress } from './range-address.util';

function normalizeColumnHint(hint: string): string {
  return hint.replace(/^the\s+/i, '').trim();
}

function extractSortColumnHint(text: string): string | undefined {
  const patterns = [
    /\bsort(?:ed)?\s+(?:the\s+)?values?\s+of\s+["']?(.+?)["']?\s+in\s+(?:ascending|descending)(?:\s+order)?/i,
    /\bsort(?:ed)?\s+(?:the\s+)?values?\s+of\s+["']?([^"'\n,]+)["']?(?:\s*$|\s+column\b)/i,
    /\bsort(?:ed)?\s+(?:the\s+)?(?:sheet\s+)?based\s+on\s+(?:the\s+)?["']?([^"'\n,]+?)["']?(?:\s+column|\s+in\b|\s*$)/i,
    /\bsort(?:ed)?\s+(?:the\s+)?(?:column\s+)["']?([^"'\n,]+?)["']?\s+in\s+(?:ascending|descending)(?:\s+order)?/i,
    /\bsort(?:ed)?\s+(?:the\s+sheet\s+)?(?:by|on)\s+(?:the\s+)?["']?([^"'\n,]+?)["']?(?:\s+column)?\s*$/i,
    /\bsort(?:ed)?\s+(?:the\s+sheet\s+)?(?:by|on)\s+(?:the\s+)?["']?([^"'\n,]+?)["']?(?:\s+column)?/i,
    /\bbased\s+on\s+(?:the\s+)?["']?([^"'\n,]+?)["']?(?:\s+column)?/i,
    /\bby\s+(?:the\s+)?["']?([A-Za-z][A-Za-z0-9 _%-]*)["']?\s*(?:column)?/i,
  ];
  for (const pattern of patterns) {
    const match = pattern.exec(text);
    if (match?.[1]) {
      return normalizeColumnHint(match[1]);
    }
  }
  return undefined;
}

function resolveColumnIndex(headers: string[], hint: string): number {
  const normalizedHint = hint.trim().toLowerCase();
  const exact = headers.findIndex((h) => h.trim().toLowerCase() === normalizedHint);
  if (exact >= 0) return exact;

  const partial = headers.findIndex((h) => {
    const lower = h.trim().toLowerCase();
    return lower.includes(normalizedHint) || normalizedHint.includes(lower);
  });
  return partial;
}

function isDescending(text: string): boolean {
  return /\b(descending|desc|highest\s+to\s+lowest|z\s*to\s*a)\b/i.test(text);
}

export function buildSortFallbackAction(
  subtask: SubTask,
  context: WorkbookContext,
): SheetActionPayload | null {
  const combined = `${subtask.description}`;
  if (!/\bsort/i.test(combined)) return null;

  const sheet = context.sheets.find((s) => s.name === subtask.targetSheet);
  if (!sheet || sheet.rowCount < 2) return null;

  const headers = (sheet.values[0] ?? []).map((cell) => String(cell ?? '').trim());
  if (!headers.some(Boolean)) return null;

  const hint = extractSortColumnHint(combined);
  if (!hint) return null;

  const key = resolveColumnIndex(headers, hint);
  if (key < 0) return null;

  const range = normalizeSortRangeAddress({
    usedRange: sheet.usedRange,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    key,
  });

  return {
    type: 'SORT_RANGE',
    sheetName: subtask.targetSheet,
    range,
    key,
    ascending: !isDescending(combined),
    hasHeaders: true,
    columnName: headers[key],
  };
}
