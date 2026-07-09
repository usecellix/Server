import { SheetActionPayload } from '../types/sheet-actions.types';

export interface ShortcutHandler {
  id: string;
  description: string;
  patterns: RegExp[];
  handler: (message: string, activeSheetName?: string) => SheetActionPayload[] | null;
}

function columnLetterToIndex(letter: string): number {
  const upper = letter.trim().toUpperCase();
  let index = 0;
  for (let i = 0; i < upper.length; i += 1) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

function parseRowRange(message: string): { rowStart: number; rowEnd: number } | null {
  const rangeMatch = message.match(/rows?\s+(\d+)\s*(?:through|to|-)\s*(\d+)/i);
  if (rangeMatch) {
    return { rowStart: parseInt(rangeMatch[1], 10), rowEnd: parseInt(rangeMatch[2], 10) };
  }

  const singleMatch = message.match(/rows?\s+(\d+)/i);
  if (singleMatch) {
    const row = parseInt(singleMatch[1], 10);
    return { rowStart: row, rowEnd: row };
  }

  return null;
}

function parseColumnRef(message: string): string | null {
  const colMatch = message.match(/col(?:umn)?\s+([A-Z]+)/i);
  return colMatch ? colMatch[1].toUpperCase() : null;
}

function parseColumnRange(message: string): { colStart: string; colEnd: string } | null {
  const rangeMatch = message.match(/col(?:umn)?s?\s+([A-Z]+)\s*(?:through|to|-)\s*([A-Z]+)/i);
  if (rangeMatch) {
    return { colStart: rangeMatch[1].toUpperCase(), colEnd: rangeMatch[2].toUpperCase() };
  }

  const col = parseColumnRef(message);
  if (col) return { colStart: col, colEnd: col };
  return null;
}

function parseZoom(message: string): number | null {
  const match = message.match(/(\d+)\s*%?/);
  return match ? parseInt(match[1], 10) : null;
}

function parseRowHeight(message: string): number | null {
  const match = message.match(/(\d+(?:\.\d+)?)\s*(?:px|pt|points?)?/i);
  return match ? parseFloat(match[1]) : null;
}

function parseColumnWidth(message: string): number | null {
  const match = message.match(/(\d+(?:\.\d+)?)\s*(?:px|pt|chars?|characters?)?/i);
  return match ? parseFloat(match[1]) : null;
}

function parseSheetName(message: string): string | null {
  const quoted = message.match(/["']([^"']+)["']/);
  if (quoted) return quoted[1];

  const named = message.match(/sheet\s+([A-Za-z0-9_\- ]+?)(?:\s*$|\s+and|\s+please)/i);
  return named ? named[1].trim() : null;
}

function parseColor(message: string): string | null {
  const colors: Record<string, string> = {
    red: '#FF0000',
    blue: '#0000FF',
    green: '#00FF00',
    yellow: '#FFFF00',
    orange: '#FFA500',
    purple: '#800080',
    pink: '#FFC0CB',
    black: '#000000',
    white: '#FFFFFF',
    gray: '#808080',
    grey: '#808080',
    teal: '#008080',
  };
  for (const [name, hex] of Object.entries(colors)) {
    if (message.toLowerCase().includes(name)) return hex;
  }
  const hexMatch = message.match(/#([0-9A-Fa-f]{3,6})/);
  return hexMatch ? hexMatch[0] : null;
}

function toRowAction(
  type: SheetActionPayload['type'],
  range: { rowStart: number; rowEnd: number },
  activeSheetName?: string,
): SheetActionPayload {
  const first = Math.min(range.rowStart, range.rowEnd);
  const last = Math.max(range.rowStart, range.rowEnd);
  return {
    type,
    sheetName: activeSheetName,
    row: first - 1,
    rowCount: last - first + 1,
  };
}

function toColumnAction(
  type: SheetActionPayload['type'],
  cols: { colStart: string; colEnd: string },
  activeSheetName?: string,
): SheetActionPayload {
  const startIdx = columnLetterToIndex(cols.colStart);
  const endIdx = columnLetterToIndex(cols.colEnd);
  return {
    type,
    sheetName: activeSheetName,
    col: Math.min(startIdx, endIdx),
    colCount: Math.abs(endIdx - startIdx) + 1,
  };
}

/** Natural-language conditions must not match literal layout shortcuts. */
export function hasConditionalShortcutBlocker(text: string): boolean {
  return /\b(where|if|that have|that has|with totals|but allow|except|only when)\b/i.test(text);
}

const SHORTCUT_REGISTRY: ShortcutHandler[] = [
  {
    id: 'freeze-top-row',
    description: 'Freeze the top row',
    patterns: [
      /freeze\s+(?:the\s+)?(?:top\s+row|first\s+row|row\s+1|header)/i,
      /lock\s+(?:the\s+)?(?:top\s+row|first\s+row|header\s+row)/i,
    ],
    handler: (_msg, activeSheetName) => [
      { type: 'FREEZE_PANES', sheetName: activeSheetName, freezeRows: 1, freezeColumns: 0 },
    ],
  },
  {
    id: 'freeze-row-count',
    description: 'Freeze first N rows',
    patterns: [
      /freeze\s+(?:the\s+)?(?:first|top)\s+(\d+)\s+rows?/i,
      /freeze\s+rows?\s+1\s*(?:to|through|-)\s*(\d+)/i,
    ],
    handler: (msg, activeSheetName) => {
      const firstN = msg.match(/freeze\s+(?:the\s+)?(?:first|top)\s+(\d+)\s+rows?/i);
      if (firstN) {
        const n = Number(firstN[1]);
        if (n >= 1 && n <= 500) {
          return [
            { type: 'FREEZE_PANES', sheetName: activeSheetName, freezeRows: n, freezeColumns: 0 },
          ];
        }
      }
      const through = msg.match(/freeze\s+rows?\s+1\s*(?:to|through|-)\s*(\d+)/i);
      if (through) {
        const n = Number(through[1]);
        if (n >= 1 && n <= 500) {
          return [
            { type: 'FREEZE_PANES', sheetName: activeSheetName, freezeRows: n, freezeColumns: 0 },
          ];
        }
      }
      return null;
    },
  },
  {
    id: 'freeze-first-column',
    description: 'Freeze the first column',
    patterns: [
      /freeze\s+(first\s+col|left\s+col|column\s+a)/i,
      /lock\s+(first\s+col|left\s+col)/i,
    ],
    handler: (_msg, activeSheetName) => [
      { type: 'FREEZE_PANES', sheetName: activeSheetName, freezeRows: 0, freezeColumns: 1 },
    ],
  },
  {
    id: 'freeze-both',
    description: 'Freeze first row and first column',
    patterns: [/freeze\s+(both|row\s+and\s+col|top\s+row\s+and\s+(first\s+)?col)/i],
    handler: (_msg, activeSheetName) => [
      { type: 'FREEZE_PANES', sheetName: activeSheetName, freezeRows: 1, freezeColumns: 1 },
    ],
  },
  {
    id: 'unfreeze',
    description: 'Remove all frozen panes',
    patterns: [/unfreeze|remove\s+freeze|clear\s+freeze|unlock\s+panes?/i],
    handler: (_msg, activeSheetName) => [{ type: 'UNFREEZE_PANES', sheetName: activeSheetName }],
  },
  {
    id: 'hide-row',
    description: 'Hide one or more rows',
    patterns: [/(?<![a-z])hide\s+row/i],
    handler: (msg, activeSheetName) => {
      const range = parseRowRange(msg);
      if (!range) return null;
      return [toRowAction('HIDE_ROW', range, activeSheetName)];
    },
  },
  {
    id: 'unhide-row',
    description: 'Unhide one or more rows',
    patterns: [/(?:unhide|show)\s+row/i],
    handler: (msg, activeSheetName) => {
      const range = parseRowRange(msg);
      if (!range) return null;
      return [toRowAction('UNHIDE_ROW', range, activeSheetName)];
    },
  },
  {
    id: 'hide-column',
    description: 'Hide one or more columns',
    patterns: [/(?<![a-z])hide\s+col/i],
    handler: (msg, activeSheetName) => {
      const cols = parseColumnRange(msg);
      if (!cols) return null;
      return [toColumnAction('HIDE_COLUMN', cols, activeSheetName)];
    },
  },
  {
    id: 'unhide-column',
    description: 'Unhide one or more columns',
    patterns: [/(?:unhide|show)\s+col/i],
    handler: (msg, activeSheetName) => {
      const cols = parseColumnRange(msg);
      if (!cols) return null;
      return [toColumnAction('UNHIDE_COLUMN', cols, activeSheetName)];
    },
  },
  {
    id: 'set-zoom',
    description: 'Set sheet zoom level',
    patterns: [/(?:zoom|scale)\s+(?:to\s+)?(\d+)\s*%?/i, /set\s+zoom\s+(?:to\s+)?(\d+)\s*%?/i],
    handler: (msg, activeSheetName) => {
      const zoomPercent = parseZoom(msg.replace(/.*(?:zoom|scale)\s+(?:to\s+)?/i, ''));
      if (!zoomPercent || zoomPercent < 10 || zoomPercent > 400) return null;
      return [{ type: 'SET_ZOOM', sheetName: activeSheetName, zoomPercent }];
    },
  },
  {
    id: 'zoom-100',
    description: 'Reset zoom to 100%',
    patterns: [/reset\s+zoom|zoom\s+(?:to\s+)?(?:normal|default|100)/i],
    handler: (_msg, activeSheetName) => [
      { type: 'SET_ZOOM', sheetName: activeSheetName, zoomPercent: 100 },
    ],
  },
  {
    id: 'protect-sheet',
    description: 'Protect the current sheet',
    patterns: [/(?<![a-z])protect\s+(?:this\s+)?sheet/i, /(?<![a-z])lock\s+(?:this\s+)?sheet/i],
    handler: (_msg, activeSheetName) => [{ type: 'PROTECT_SHEET', sheetName: activeSheetName }],
  },
  {
    id: 'unprotect-sheet',
    description: 'Unprotect the current sheet',
    patterns: [/unprotect\s+(?:this\s+)?sheet/i, /unlock\s+(?:this\s+)?sheet/i],
    handler: (_msg, activeSheetName) => [{ type: 'UNPROTECT_SHEET', sheetName: activeSheetName }],
  },
  {
    id: 'set-row-height',
    description: 'Set row height for one or more rows',
    patterns: [/set\s+row\s+height/i, /row\s+height\s+to/i, /make\s+rows?\s+taller|shorter/i],
    handler: (msg, activeSheetName) => {
      const range = parseRowRange(msg);
      const height = parseRowHeight(msg.replace(/.*(?:height|to|taller|shorter)\s*/i, ''));
      if (!height) return null;
      const rowStart = range?.rowStart ?? 1;
      const rowEnd = range?.rowEnd ?? rowStart;
      return [
        {
          type: 'SET_ROW_HEIGHT',
          sheetName: activeSheetName,
          row: rowStart - 1,
          rowCount: rowEnd - rowStart + 1,
          height,
        },
      ];
    },
  },
  {
    id: 'set-column-width',
    description: 'Set column width',
    patterns: [/set\s+col(?:umn)?\s+width/i, /col(?:umn)?\s+width\s+to/i],
    handler: (msg, activeSheetName) => {
      const cols = parseColumnRange(msg);
      const width = parseColumnWidth(msg.replace(/.*(?:width|to)\s*/i, ''));
      if (!width) return null;
      const colStart = cols?.colStart ?? 'A';
      const colEnd = cols?.colEnd ?? colStart;
      const colAction = toColumnAction('SET_COLUMN_WIDTH', { colStart, colEnd }, activeSheetName);
      return [{ ...colAction, width }];
    },
  },
  {
    id: 'autofit-columns',
    description: 'Auto-fit all column widths',
    patterns: [
      /auto\s*fit\s+col/i,
      /fit\s+col(?:umn)?s?\s+(?:to\s+)?(?:content|data|text)/i,
      /resize\s+col(?:umn)?s?\s+(?:to\s+)?(?:fit|content)/i,
      /adjust\s+col(?:umn)?\s+width/i,
    ],
    handler: (_msg, activeSheetName) => [{ type: 'AUTOFIT_COLUMNS', sheetName: activeSheetName }],
  },
  {
    id: 'hide-sheet',
    description: 'Hide a sheet tab',
    patterns: [/(?<![a-z])hide\s+(?:the\s+)?(?:sheet|tab)/i],
    handler: (msg, activeSheetName) => {
      const name = parseSheetName(msg);
      return [{ type: 'HIDE_SHEET', sheetName: name ?? activeSheetName }];
    },
  },
  {
    id: 'show-sheet',
    description: 'Show a hidden sheet tab',
    patterns: [/(?:show|unhide)\s+(?:the\s+)?(?:sheet|tab)/i],
    handler: (msg, activeSheetName) => {
      const name = parseSheetName(msg);
      return [{ type: 'SHOW_SHEET', sheetName: name ?? activeSheetName }];
    },
  },
  {
    id: 'set-sheet-color',
    description: 'Change sheet tab color',
    patterns: [
      /(?:color|colour)\s+(?:this\s+)?(?:sheet|tab)/i,
      /(?:set|change)\s+(?:sheet|tab)\s+(?:color|colour)/i,
      /make\s+(?:the\s+)?(?:sheet|tab)\s+(\w+)/i,
    ],
    handler: (msg, activeSheetName) => {
      const color = parseColor(msg);
      if (!color) return null;
      return [{ type: 'SET_SHEET_COLOR', sheetName: activeSheetName, color }];
    },
  },
  {
    id: 'delete-comment',
    description: 'Delete a comment from a cell',
    patterns: [/(?:delete|remove|clear)\s+comment/i],
    handler: (msg, activeSheetName) => {
      const cellMatch = msg.match(/\b([A-Z]+\d+)\b/i);
      return [
        {
          type: 'DELETE_COMMENT',
          sheetName: activeSheetName,
          address: cellMatch ? cellMatch[1].toUpperCase() : undefined,
        },
      ];
    },
  },
];

export function routeShortcutAction(
  message: string,
  activeSheetName?: string,
): SheetActionPayload[] | null {
  const text = message.trim();
  if (!text) return null;
  if (hasConditionalShortcutBlocker(text)) return null;

  for (const shortcut of SHORTCUT_REGISTRY) {
    const matches = shortcut.patterns.some((pattern) => pattern.test(text));
    if (matches) {
      const result = shortcut.handler(text, activeSheetName);
      if (result?.length) return result;
    }
  }

  return null;
}

export function listShortcuts(): Array<{ id: string; description: string }> {
  return SHORTCUT_REGISTRY.map(({ id, description }) => ({ id, description }));
}

function rowDisplayRange(action: SheetActionPayload): { start: number; end: number } {
  const start =
    (action as SheetActionPayload & { rowStart?: number }).rowStart ??
    (typeof action.row === 'number' ? action.row + 1 : 1);
  const end =
    (action as SheetActionPayload & { rowEnd?: number }).rowEnd ??
    (typeof action.row === 'number' && action.rowCount
      ? action.row + action.rowCount - 1
      : start);
  return { start, end };
}

export function buildShortcutAnswer(actions: SheetActionPayload[]): string {
  if (!actions || actions.length === 0) return 'Done.';

  const descriptions: Record<string, (action: SheetActionPayload) => string> = {
    FREEZE_PANES: (action) =>
      action.freezeRows && action.freezeColumns
        ? 'Frozen the top row and first column.'
        : action.freezeRows
          ? `Frozen the top ${action.freezeRows === 1 ? 'row' : `${action.freezeRows} rows`}.`
          : `Frozen the first ${action.freezeColumns === 1 ? 'column' : `${action.freezeColumns} columns`}.`,
    UNFREEZE_PANES: () => 'Removed all frozen panes.',
    HIDE_ROW: (action) => {
      const { start, end } = rowDisplayRange(action);
      return start === end ? `Hidden row ${start}.` : `Hidden rows ${start}–${end}.`;
    },
    UNHIDE_ROW: (action) => {
      const { start, end } = rowDisplayRange(action);
      return start === end ? `Shown row ${start}.` : `Shown rows ${start}–${end}.`;
    },
    HIDE_COLUMN: (action) => {
      const start =
        (action as SheetActionPayload & { columnStart?: string }).columnStart ??
        (typeof action.col === 'number' ? String.fromCharCode(65 + action.col) : 'A');
      const end =
        (action as SheetActionPayload & { columnEnd?: string }).columnEnd ??
        (typeof action.col === 'number' && action.colCount
          ? String.fromCharCode(65 + action.col + action.colCount - 1)
          : start);
      return start === end ? `Hidden column ${start}.` : `Hidden columns ${start}–${end}.`;
    },
    UNHIDE_COLUMN: (action) => {
      const start =
        (action as SheetActionPayload & { columnStart?: string }).columnStart ??
        (typeof action.col === 'number' ? String.fromCharCode(65 + action.col) : 'A');
      const end =
        (action as SheetActionPayload & { columnEnd?: string }).columnEnd ??
        (typeof action.col === 'number' && action.colCount
          ? String.fromCharCode(65 + action.col + action.colCount - 1)
          : start);
      return start === end ? `Shown column ${start}.` : `Shown columns ${start}–${end}.`;
    },
    SET_ZOOM: (action) => {
      const zoom =
        (action as SheetActionPayload & { zoom?: number }).zoom ?? action.zoomPercent ?? 100;
      return `Zoom set to ${zoom}%.`;
    },
    PROTECT_SHEET: () => 'Sheet is now protected.',
    UNPROTECT_SHEET: () => 'Sheet protection removed.',
    SET_ROW_HEIGHT: (action) => `Row height set to ${action.height}.`,
    SET_COLUMN_WIDTH: (action) => `Column width set to ${action.width}.`,
    AUTOFIT_COLUMNS: () => 'Column widths adjusted to fit content.',
    HIDE_SHEET: (action) => `Sheet "${action.sheetName ?? 'current'}" is now hidden.`,
    SHOW_SHEET: (action) => `Sheet "${action.sheetName ?? 'current'}" is now visible.`,
    SET_SHEET_COLOR: (action) => `Tab color changed to ${action.color}.`,
    ADD_COMMENT: (action) => `Comment added to ${action.address ?? 'selected cell'}.`,
    DELETE_COMMENT: (action) => `Comment removed from ${action.address ?? 'selected cell'}.`,
  };

  const first = actions[0];
  const describe = descriptions[first.type];
  return describe ? describe(first) : `Applied ${first.type.toLowerCase().replace(/_/g, ' ')}.`;
}
