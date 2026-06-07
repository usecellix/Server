import { SheetActionPayload } from '../types/sheet-actions.types';

export interface TablePlan {
  headers: string[];
  rows: unknown[][];
  rowCount: number;
}

const STATUS_VALUES = ['Active', 'Pending', 'Closed', 'Review', 'Hold'];

export function parseTableCreateRequest(message: string): TablePlan | null {
  const lower = message.toLowerCase();
  const isCreate =
    /\b(create|generate|add|populate|fill|make|build|give|insert)\b/.test(lower) &&
    (/\b(row|rows|header|headers|column|columns|table|sheet|dummy|sample)\b/.test(lower) ||
      /\bs\.?\s*no\b/.test(lower));

  if (!isCreate) return null;

  const headers = extractHeaders(message);
  if (headers.length < 2) return null;

  const rowCount = extractRowCount(message) ?? 5;
  const rows = buildDummyRows(headers, rowCount);

  return { headers, rows, rowCount };
}

export function buildWriteTableAction(plan: TablePlan): SheetActionPayload {
  return {
    type: 'WRITE_TABLE',
    headers: plan.headers,
    rows: plan.rows,
  };
}

export function buildTableActionsFromMessage(message: string): SheetActionPayload[] | null {
  const plan = parseTableCreateRequest(message);
  if (!plan) return null;
  return [buildWriteTableAction(plan)];
}

function extractRowCount(message: string): number | null {
  const patterns = [
    /\b(\d{1,3})\s*rows?\b/i,
    /\bcreate\s+(\d{1,3})\b/i,
    /\bgive\s+(\d{1,3})\b/i,
  ];
  for (const pattern of patterns) {
    const match = pattern.exec(message);
    if (match) {
      const n = Number.parseInt(match[1], 10);
      if (n >= 1 && n <= 500) return n;
    }
  }
  return null;
}

function extractHeaders(message: string): string[] {
  const patterns = [
    /\bwith\s+headers?\s+(.+?)(?=\s*,?\s*(?:give|and\s+give|for\s+this)\b|$)/i,
    /\bheaders?\s*:\s*(.+?)(?=\s*,?\s*(?:give|for)\b|$)/i,
    /\bcolumns?\s+(?:as\s+)?(.+?)(?=\s*,?\s*(?:give|for)\b|$)/i,
  ];

  for (const pattern of patterns) {
    const match = pattern.exec(message);
    if (!match) continue;
    const parsed = splitHeaderList(match[1]);
    if (parsed.length >= 2) return parsed;
  }

  return [];
}

function splitHeaderList(raw: string): string[] {
  const trimmed = raw.split(/\bgive\b|\bwith\b|\bfor\b/i)[0].trim();
  return trimmed
    .split(/,|\band\b/gi)
    .map((part) => part.trim().replace(/^["']|["']$/g, ''))
    .filter((part) => part.length > 0 && part.length < 40 && !/^\d+\s*(rows?|dummy)/i.test(part));
}

function buildDummyRows(headers: string[], rowCount: number): unknown[][] {
  return Array.from({ length: rowCount }, (_, rowIndex) =>
    headers.map((header) => dummyCellValue(header, rowIndex)),
  );
}

function dummyCellValue(header: string, rowIndex: number): unknown {
  const h = header.toLowerCase();

  if (/\bs\.?\s*no|sr\.?\s*no|serial|#/i.test(h)) {
    return rowIndex + 1;
  }
  if (/tax\s*%|tax percent|gst\s*%/i.test(h)) {
    return [0, 5, 12, 18, 28][rowIndex % 5];
  }
  if (/amount|value|price|total|rupee|₹/i.test(h)) {
    return (rowIndex + 1) * 12500 + rowIndex * 3500;
  }
  if (/return\s*rate|rate\s*%|rate/i.test(h)) {
    return [2.5, 5, 7.5, 10, 12.5][rowIndex % 5];
  }
  if (/status|state/i.test(h)) {
    return STATUS_VALUES[rowIndex % STATUS_VALUES.length];
  }
  if (/date/i.test(h)) {
    const day = String(rowIndex + 1).padStart(2, '0');
    return `${day}-04-2024`;
  }
  if (/gstin/i.test(h)) {
    return `29AABCT${1000 + rowIndex}L1Z5`;
  }
  if (/name|supplier|party/i.test(h)) {
    return `Sample Vendor ${rowIndex + 1}`;
  }

  return `Value ${rowIndex + 1}`;
}

export function actionsNeedTableFallback(
  message: string,
  actions: SheetActionPayload[] | undefined,
): boolean {
  const plan = parseTableCreateRequest(message);
  if (!plan) return false;
  if (!actions?.length) return true;

  const hasWriteTable = actions.some((a) => a.type === 'WRITE_TABLE');
  if (hasWriteTable) return false;

  const addRowCount = actions.filter((a) => a.type === 'ADD_ROW').length;
  const headerCells = actions.filter((a) => a.type === 'SET_CELL' && a.row === 0).length;

  return addRowCount < plan.rowCount || headerCells < plan.headers.length;
}
