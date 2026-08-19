import { WorkbookContext } from '../../agents/types/agent.types';
import { SheetAction, RangeFilterSpec } from '../types/sheet-actions.types';

const LIGHT_RED = '#FFC7CE';

const COLOR_ALIASES: Record<string, string> = {
  'light red': LIGHT_RED,
  lightred: LIGHT_RED,
  red: '#FF6B6B',
  'light yellow': '#FFF2CC',
  yellow: '#FFE699',
  'light green': '#C6EFCE',
  green: '#A9D08E',
  'light blue': '#BDD7EE',
  blue: '#9BC2E6',
  orange: '#FCE4D6',
  pink: '#F8CBAD',
};

function stripSheetPrefix(range: string): string {
  const bang = range.lastIndexOf('!');
  return bang >= 0 ? range.slice(bang + 1).replace(/^'|'$/g, '') : range;
}

function mapConditionOperator(raw: unknown): RangeFilterSpec['operator'] {
  const text = String(raw ?? '')
    .trim()
    .toUpperCase()
    .replace(/[\s-]+/g, '_');
  if (text === 'TEXT_EQ' || text === 'EQ' || text === 'EQUALS' || text === 'EQUAL') {
    return 'equals';
  }
  if (text === 'TEXT_NE' || text === 'NE' || text === 'NOT_EQUALS' || text === 'NOT_EQUAL') {
    return 'notEquals';
  }
  if (text === 'CONTAINS' || text === 'TEXT_CONTAINS') return 'contains';
  if (text === 'GT' || text === 'GREATER_THAN') return 'greaterThan';
  if (text === 'LT' || text === 'LESS_THAN') return 'lessThan';
  return 'equals';
}

function resolveColumnName(
  column: unknown,
  headers: string[],
): string | null {
  if (typeof column === 'string' && column.trim()) {
    const named = headers.find(
      (h) => h.trim().toLowerCase() === column.trim().toLowerCase(),
    );
    return named ?? column.trim();
  }
  if (typeof column === 'number' && Number.isFinite(column)) {
    // Prefer 1-based Excel ordinal within the header row (K=11 → Payment Status).
    if (Number.isInteger(column) && column >= 1 && column <= headers.length) {
      return headers[column - 1];
    }
    if (Number.isInteger(column) && column >= 0 && column < headers.length) {
      return headers[column];
    }
  }
  return null;
}

function resolveFillColor(action: Record<string, unknown>): string | undefined {
  const format =
    action.format && typeof action.format === 'object'
      ? (action.format as Record<string, unknown>)
      : undefined;
  const raw =
    (typeof action.color === 'string' && action.color) ||
    (typeof format?.fillColor === 'string' && format.fillColor) ||
    (typeof action.fillColor === 'string' && action.fillColor) ||
    undefined;
  if (!raw) return undefined;
  const alias = COLOR_ALIASES[raw.trim().toLowerCase()];
  return alias ?? raw;
}

function wantsClearFill(record: Record<string, unknown>): boolean {
  const format =
    record.format && typeof record.format === 'object'
      ? (record.format as Record<string, unknown>)
      : undefined;
  if (format?.clearFill === true) return true;
  const fill = typeof format?.fillColor === 'string' ? format.fillColor.trim().toLowerCase() : '';
  return fill === 'none' || fill === 'clear' || fill === 'transparent';
}

export function isClearHighlightMessage(message: string): boolean {
  const normalized = message.toLowerCase().replace(/\bremvoe\b/g, 'remove');
  return (
    /\b(remove|clear|unhighlight)\b.*\b(highlights?|fills?|colou?rs?)\b/i.test(normalized) ||
    /\b(highlights?|fills?|colou?rs?)\b.*\b(remove|clear)\b/i.test(normalized)
  );
}

/**
 * Spec 24: header-row cosmetics ("highlight the header row") — not a data-row match.
 */
export function isHeaderRowFormatRequest(message: string): boolean {
  const lower = message.toLowerCase();
  if (!/\b(header|headers)\b/.test(lower)) return false;
  // Conditional highlight of rows that mention "header" should not match; need format intent.
  return (
    /\b(header|headers)\b.*\b(row|bg|background|fill|highlight|bold|colou?r|green|red|yellow|blue)\b/i.test(
      lower,
    ) ||
    /\b(highlight|fill|bg|background|bold|colou?r)\b.*\b(header|headers)\b/i.test(lower)
  );
}

function resolveColorFromMessage(message: string): string | undefined {
  const lower = message.toLowerCase();
  // "light bg green" / "light background red" (common chat phrasing)
  const lightLoose = /\blight\s+(?:bg|background|fill)?\s*(red|green|yellow|blue)\b/i.exec(lower);
  if (lightLoose) {
    const key = `light ${lightLoose[1]!.toLowerCase()}`;
    if (COLOR_ALIASES[key]) return COLOR_ALIASES[key];
  }
  // Prefer multi-word aliases first
  const ordered = Object.keys(COLOR_ALIASES).sort((a, b) => b.length - a.length);
  for (const key of ordered) {
    if (lower.includes(key)) return COLOR_ALIASES[key];
  }
  return undefined;
}

function wantsHeaderBold(message: string): boolean {
  return /\bbold\b/i.test(message);
}

function resolveHeaderColCount(workbookContext: WorkbookContext, sheetName: string): number {
  const sheet = workbookContext.sheets.find((s) => s.name === sheetName);
  if (!sheet) return workbookContext.sheets[0]?.columnCount || 1;
  const fromValues = Math.max(
    sheet.columnCount || 0,
    ...(sheet.values ?? []).map((row) => (Array.isArray(row) ? row.length : 0)),
    0,
  );
  const used = stripSheetPrefix(sheet.usedRange ?? 'A1');
  const rangeMatch = /^[A-Za-z]+(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(used);
  if (rangeMatch?.[2]) {
    const endCol = rangeMatch[2];
    let col = 0;
    for (const ch of endCol.toUpperCase()) {
      col = col * 26 + (ch.charCodeAt(0) - 64);
    }
    return Math.max(fromValues, col, 1);
  }
  return Math.max(fromValues, 1);
}

/** Deterministic FORMAT_RANGE for the full header row. */
export function buildHeaderRowFormatAction(
  workbookContext: WorkbookContext,
  userMessage: string,
  seed?: SheetAction,
): SheetAction {
  const record = (seed ?? {}) as unknown as Record<string, unknown>;
  const sheetName =
    (typeof record.sheetName === 'string' && record.sheetName) ||
    workbookContext.activeSheetName;
  const colCount = resolveHeaderColCount(workbookContext, sheetName);
  const clearFill = isClearHighlightMessage(userMessage) || wantsClearFill(record);
  const fillColor =
    resolveFillColor(record) ?? resolveColorFromMessage(userMessage) ?? LIGHT_RED;
  const bold = wantsHeaderBold(userMessage);

  if (clearFill) {
    return {
      type: 'FORMAT_RANGE',
      sheetName,
      row: 0,
      col: 0,
      rowCount: 1,
      colCount,
      format: { clearFill: true },
    };
  }

  const format: NonNullable<SheetAction['format']> = { fillColor };
  if (bold) format.bold = true;

  return {
    type: 'FORMAT_RANGE',
    sheetName,
    row: 0,
    col: 0,
    rowCount: 1,
    colCount,
    format,
  };
}

function resolveUsedRange(workbookContext: WorkbookContext, sheetName: string): string {
  const sheet = workbookContext.sheets.find((s) => s.name === sheetName);
  const used = sheet?.usedRange ?? workbookContext.sheets[0]?.usedRange ?? 'A1';
  return stripSheetPrefix(used);
}

function buildFilterFromCondition(
  condition: Record<string, unknown>,
  headers: string[],
): RangeFilterSpec | null {
  const column = resolveColumnName(condition.column ?? condition.col, headers);
  if (!column) return null;
  if (typeof condition.value !== 'string' && typeof condition.value !== 'number') {
    return null;
  }
  return {
    column,
    operator: mapConditionOperator(condition.type ?? condition.operator),
    value: condition.value,
  };
}

/**
 * Normalize Tier-1 / executor output into FORMAT_MATCHING_ROWS when the model
 * invents a condition on HIGHLIGHT_CELL / FORMAT_RANGE, or emit a clean
 * FORMAT_MATCHING_ROWS action.
 */
export function normalizeFormatMatchingRowsAction(
  action: SheetAction,
  workbookContext: WorkbookContext,
  userMessage?: string,
): SheetAction {
  const record = action as unknown as Record<string, unknown>;
  const sheetName =
    (typeof record.sheetName === 'string' && record.sheetName) ||
    workbookContext.activeSheetName;

  const sheet = workbookContext.sheets.find((s) => s.name === sheetName);
  const headers = (sheet?.values?.[0] ?? []).map((cell) => String(cell ?? '').trim());

  if (action.type === 'FORMAT_MATCHING_ROWS') {
    const filter =
      record.filter && typeof record.filter === 'object'
        ? (record.filter as Record<string, unknown>)
        : undefined;
    const column =
      filter && resolveColumnName(filter.column, headers);
    const operator =
      filter && typeof filter.operator === 'string'
        ? (filter.operator as RangeFilterSpec['operator'])
        : 'equals';
    const value = filter?.value;
    if (
      !column ||
      (typeof value !== 'string' && typeof value !== 'number')
    ) {
      return action;
    }
    const range =
      (typeof record.range === 'string' && stripSheetPrefix(record.range)) ||
      resolveUsedRange(workbookContext, sheetName);
    const clearFill = wantsClearFill(record) || (userMessage ? isClearHighlightMessage(userMessage) : false);
    const format =
      record.format && typeof record.format === 'object'
        ? (record.format as SheetAction['format'])
        : clearFill
          ? { clearFill: true }
          : { fillColor: LIGHT_RED };
    return {
      type: 'FORMAT_MATCHING_ROWS',
      sheetName,
      range,
      hasHeaders: record.hasHeaders !== false,
      filter: { column, operator, value },
      format: clearFill
        ? { clearFill: true }
        : {
            ...format,
            fillColor: format?.fillColor ?? resolveFillColor(record) ?? LIGHT_RED,
          },
    };
  }

  const condition =
    record.condition && typeof record.condition === 'object'
      ? (record.condition as Record<string, unknown>)
      : undefined;

  if (
    condition &&
    (action.type === 'HIGHLIGHT_CELL' || action.type === 'FORMAT_RANGE')
  ) {
    const filter = buildFilterFromCondition(condition, headers);
    if (!filter) return action;
    const clearFill = wantsClearFill(record) || (userMessage ? isClearHighlightMessage(userMessage) : false);
    return {
      type: 'FORMAT_MATCHING_ROWS',
      sheetName,
      range: resolveUsedRange(workbookContext, sheetName),
      hasHeaders: true,
      filter,
      format: clearFill
        ? { clearFill: true }
        : { fillColor: resolveFillColor(record) ?? LIGHT_RED },
    };
  }

  return action;
}

export function normalizeTier1ConditionalFormatActions(
  actions: SheetAction[],
  workbookContext: WorkbookContext,
  userMessage?: string,
): SheetAction[] {
  if (userMessage && isHeaderRowFormatRequest(userMessage)) {
    const seed = actions[0];
    return [buildHeaderRowFormatAction(workbookContext, userMessage, seed)];
  }
  return actions.map((action) =>
    normalizeFormatMatchingRowsAction(action, workbookContext, userMessage),
  );
}

/** Normalize any Tier-1 header-format batch (HEADER_FORMAT hint or misrouted match). */
export function normalizeTier1HeaderFormatActions(
  actions: SheetAction[],
  workbookContext: WorkbookContext,
  userMessage: string,
): SheetAction[] {
  return [buildHeaderRowFormatAction(workbookContext, userMessage, actions[0])];
}
