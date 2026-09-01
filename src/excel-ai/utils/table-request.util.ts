import { SheetActionPayload } from '../types/sheet-actions.types';
import {
  detectCompoundSheetFollowUp,
  extractSheetNameFromPrompt,
  sanitizeExcelSheetName,
} from './sheet-name.util';
import { hasCompoundSignals } from './complexity-classifier.util';

export interface TablePlan {
  headers: string[];
  rows: unknown[][];
  rowCount: number;
}

const STATUS_VALUES = ['Active', 'Pending', 'Closed', 'Review', 'Hold'];

export const DEFAULT_GST_HEADERS = [
  'S.No',
  'Invoice Date',
  'Supplier GSTIN',
  'Supplier Name',
  'Invoice Amount',
  'IGST',
  'CGST',
  'SGST',
  'Narration',
];

export function detectCreateNewSheetIntent(message: string): boolean {
  return /\b(create|add)\s+(?:an?\s+)?(?:(?:new|empty|blank)\s+)*sheet/i.test(message);
}

/** Prompts that need LLM planning (data, copy, sort, etc.) — not empty-sheet-only. */
export function detectSheetDataGenerationIntent(message: string): boolean {
  const lower = message.toLowerCase();
  if (detectCompoundSheetFollowUp(message)) return true;
  if (/\b(as\s+a\s+copy|copy\s+of|duplicate|clone)\b/i.test(message)) return true;
  if (
    /\bsort(?:\s+the\s+values?\s+of|\s+(?:the\s+)?(?:sheet\s+)?(?:based\s+on|by|on)|\s+based\s+on|\s+by|\s+on|\s+column\b)/i.test(
      message,
    ) ||
    /\bin\s+(?:ascending|descending)\s+order\b/i.test(message)
  ) {
    return true;
  }

  const hasDataKeyword =
    /\b(dummy|sample|data|values?|rows?|headers?|columns?|populate|generate|fill|table|content|gst)\b/i.test(
      lower,
    );
  const hasCreateKeyword = /\b(create|add|generate|populate|fill|make|build|give|insert)\b/i.test(
    lower,
  );

  if (/\bchart\b/i.test(message) && /\banaly(?:sis|ze|se)\b/i.test(message)) return true;

  return hasDataKeyword && hasCreateKeyword;
}

export function parseTableCreateRequest(message: string): TablePlan | null {
  const lower = message.toLowerCase();
  const isCreate =
    /\b(create|generate|add|populate|fill|make|build|give|insert)\b/.test(lower) &&
    (/\b(row|rows|header|headers|column|columns|table|sheet|dummy|sample|data)\b/.test(lower) ||
      /\bs\.?\s*no\b/.test(lower));

  if (!isCreate) return null;

  let headers = extractHeaders(message);
  if (headers.length < 2 && /\bgst\b/i.test(message)) {
    headers = [...DEFAULT_GST_HEADERS];
  }
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

export interface DeterministicTableResult {
  plan: TablePlan;
  actions: SheetActionPayload[];
}

/**
 * The single gate for "should this message deterministically produce a WRITE_TABLE
 * action, with no LLM/planner involved." conversation.service.ts calls this from
 * two different moments — before the LLM is asked anything, and again as a
 * last-resort fallback when the LLM's response failed to parse — and both used
 * to apply their own, different conditions. The fallback path was missing the
 * compound-signal gate entirely, so a prompt like "create a table with columns
 * Name, Age, City, and then add a chart" could still silently drop the chart
 * clause if the LLM call happened to fail to parse. One gate, called from both
 * places, means a fix here can't go stale in the other.
 */
export function tryDeterministicTableCreate(message: string): DeterministicTableResult | null {
  if (hasCompoundSignals(message)) return null;

  const isNewSheetWithData =
    detectCreateNewSheetIntent(message) && detectSheetDataGenerationIntent(message);
  if (isNewSheetWithData) return null;

  const plan = parseTableCreateRequest(message);
  if (!plan || plan.headers.length < 2) return null;

  const actions = buildTableActionsFromMessage(message);
  if (!actions?.length) return null;

  return { plan, actions };
}

function extractRowCount(message: string): number | null {
  const patterns = [
    /\b(\d{1,3})\s+dummy\b/i,
    /\bwith\s+(\d{1,3})\s+(?:dummy|sample)\b/i,
    /\b(\d{1,3})\s+(?:dummy|sample)\s+rows?\b/i,
    /\band\s+(\d{1,3})\s+(?:dummy|sample)\s+rows?\b/i,
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

export { extractSheetNameFromPrompt } from './sheet-name.util';

/**
 * "Create a new sheet named Cellix with 5 dummy values" → empty sheet + WRITE_TABLE.
 */
export function parseNewSheetWithDummyData(message: string): {
  sheetName: string;
  headers: string[];
  rows: unknown[][];
} | null {
  const lower = message.toLowerCase();
  if (!/\b(create|add)\s+(?:an?\s+)?(?:(?:new|empty|blank)\s+)*sheet/i.test(lower)) return null;
  if (!/\b(dummy|sample|values?|data|row)\b/i.test(lower)) return null;

  const explicitName = extractSheetNameFromPrompt(message);
  const sheetName = sanitizeExcelSheetName(explicitName ?? 'New Sheet');

  let headers = extractHeaders(message);
  if (headers.length < 2) {
    headers = ['Item', 'Value', 'Status'];
  }

  const rowCount = extractRowCount(message) ?? 5;
  const rows = buildDummyRows(headers, rowCount);

  return { sheetName, headers, rows };
}

export function buildNewSheetWithDummyDataActions(message: string): SheetActionPayload[] | null {
  const plan = parseNewSheetWithDummyData(message);
  if (!plan) return null;

  return [
    { type: 'ADD_SHEET', name: sanitizeExcelSheetName(plan.sheetName) },
    {
      type: 'WRITE_TABLE',
      sheetName: plan.sheetName,
      headers: plan.headers,
      rows: plan.rows,
    },
  ];
}

function extractHeaders(message: string): string[] {
  const patterns = [
    // "add headers Job Title, Company, … and 3 sample rows"
    // Negative lookahead on "row" excludes unrelated phrases like "freeze the header row".
    /\b(?:add|set|insert|create)\s+headers?\s+(?!row\b)(.+?)(?=\s+and\s+\d|\s*,?\s*(?:give|and\s+give|for\s+this)\b|$)/i,
    /\bwith\s+headers?\s+(?!row\b)(.+?)(?=\s+and\s+\d|\s*,?\s*(?:give|and\s+give|for\s+this)\b|$)/i,
    /\bheaders?\s*:\s*(.+?)(?=\s+and\s+\d|\s*,?\s*(?:give|for)\b|$)/i,
    /\bheaders?\s+(?!row\b)(.+?)(?=\s+and\s+\d|\s*,?\s*(?:give|and\s+give|for\s+this)\b|$)/i,
    /\bcolumns?\s+(?:named|called)\s+(.+?)(?=\s+and\s+\d|\s*,?\s*(?:give|for)\b|$)/i,
    /\bcolumns?\s+(?:as\s+)?(.+?)(?=\s+and\s+\d|\s*,?\s*(?:give|for)\b|$)/i,
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
  const trimmed = raw
    .split(/\bgive\b|\bwith\b|\bfor\b|\band\s+\d/i)[0]
    .trim();
  return trimmed
    .split(/,|\band\b/gi)
    .map((part) =>
      part
        .trim()
        .replace(/^["']|["']$/g, '')
        .replace(/^(?:named|called)\s+/i, '')
        .trim(),
    )
    .filter(
      (part) =>
        part.length > 0 &&
        part.length < 40 &&
        !/^\d+\s*(rows?|dummy|sample)/i.test(part) &&
        !/^(?:dummy|sample)\s+rows?$/i.test(part),
    );
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
  if (/\bigst\b/i.test(h)) {
    const base = (rowIndex + 1) * 12500 + rowIndex * 3500;
    return Math.round(base * 0.18);
  }
  if (/\bcgst\b|\bsgst\b/i.test(h)) {
    const base = (rowIndex + 1) * 12500 + rowIndex * 3500;
    return Math.round(base * 0.09);
  }
  if (/amount|value|price|total|rupee|₹|invoice/i.test(h)) {
    return (rowIndex + 1) * 12500 + rowIndex * 3500;
  }
  if (/return\s*rate|rate\s*%|rate/i.test(h)) {
    return [2.5, 5, 7.5, 10, 12.5][rowIndex % 5];
  }
  if (/status|state/i.test(h)) {
    return STATUS_VALUES[rowIndex % STATUS_VALUES.length];
  }
  if (/email|e-mail|mail\b/i.test(h)) {
    return `student${rowIndex + 1}@example.com`;
  }
  if (/job\s*title|title|role|designation/i.test(h)) {
    return ['Sales Associate', 'Field Technician', 'Service Manager', 'Analyst', 'Intern'][
      rowIndex % 5
    ];
  }
  if (/company|shop|employer|organization|org\b/i.test(h)) {
    return `Company ${rowIndex + 1}`;
  }
  if (/date/i.test(h)) {
    const day = String(rowIndex + 1).padStart(2, '0');
    return `${day}-04-2024`;
  }
  if (/gstin/i.test(h)) {
    return `29AABCT${1000 + rowIndex}L1Z5`;
  }
  if (/name|supplier|party|student/i.test(h)) {
    return `Sample Person ${rowIndex + 1}`;
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
