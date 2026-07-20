import { Injectable } from '@nestjs/common';
import { WorkbookContext } from '../../agents/types/agent.types';
import { FormatSpec, SheetAction } from '../types/sheet-actions.types';
import { extractTier0PatternMatch } from '../utils/complexity-classifier.util';

export interface Tier0Result {
  actions: SheetAction[];
  skippedLLM: true;
}

function columnLetterToIndex(letter: string): number {
  const upper = letter.trim().toUpperCase();
  let index = 0;
  for (let i = 0; i < upper.length; i += 1) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

function parseCellReference(message: string): { row: number; col: number; rowCount: number; colCount: number } | null {
  const rangeMatch = message.match(/\b([A-Z]+)(\d+)\s*:\s*([A-Z]+)(\d+)\b/i);
  if (rangeMatch) {
    const startCol = columnLetterToIndex(rangeMatch[1]);
    const startRow = parseInt(rangeMatch[2], 10) - 1;
    const endCol = columnLetterToIndex(rangeMatch[3]);
    const endRow = parseInt(rangeMatch[4], 10) - 1;
    return {
      row: Math.min(startRow, endRow),
      col: Math.min(startCol, endCol),
      rowCount: Math.abs(endRow - startRow) + 1,
      colCount: Math.abs(endCol - startCol) + 1,
    };
  }

  const singleMatch = message.match(/\b([A-Z]+)(\d+)\b/i);
  if (!singleMatch) {
    return null;
  }

  const row = parseInt(singleMatch[2], 10) - 1;
  const col = columnLetterToIndex(singleMatch[1]);
  return { row, col, rowCount: 1, colCount: 1 };
}

function parseRowReference(message: string): { row: number; rowCount: number } | null {
  const rangeMatch = message.match(/rows?\s+(\d+)\s*(?:through|to|-)\s*(\d+)/i);
  if (rangeMatch) {
    const start = parseInt(rangeMatch[1], 10);
    const end = parseInt(rangeMatch[2], 10);
    return {
      row: Math.min(start, end) - 1,
      rowCount: Math.abs(end - start) + 1,
    };
  }

  const singleMatch = message.match(/rows?\s+(\d+)/i);
  if (singleMatch) {
    const row = parseInt(singleMatch[1], 10);
    return { row: row - 1, rowCount: 1 };
  }

  return null;
}

function parseColumnReference(message: string): { col: number; colCount: number } | null {
  const rangeMatch = message.match(/col(?:umn)?s?\s+([A-Z]+)\s*(?:through|to|-)\s*([A-Z]+)/i);
  if (rangeMatch) {
    const start = columnLetterToIndex(rangeMatch[1]);
    const end = columnLetterToIndex(rangeMatch[2]);
    return {
      col: Math.min(start, end),
      colCount: Math.abs(end - start) + 1,
    };
  }

  const singleMatch = message.match(/col(?:umn)?\s+([A-Z]+)/i);
  if (singleMatch) {
    return { col: columnLetterToIndex(singleMatch[1]), colCount: 1 };
  }

  return null;
}

function parseSheetName(message: string): string | null {
  const quoted = message.match(/["']([^"']+)["']/);
  if (quoted) {
    return quoted[1];
  }

  const named = message.match(/sheet\s+([A-Za-z0-9_\- ]+?)(?:\s*$|\s+and|\s+please)/i);
  return named ? named[1].trim() : null;
}

@Injectable()
export class Tier0DirectService {
  resolve(
    actionHint: string,
    message: string,
    workbookContext: WorkbookContext,
  ): Tier0Result | null {
    const captures = extractTier0PatternMatch(message, actionHint);
    if (!captures) {
      return null;
    }

    const activeSheet = workbookContext.activeSheetName;

    switch (actionHint) {
      case 'CELL_FORMAT':
        return this.resolveCellFormat(message, captures, activeSheet);
      case 'FREEZE_PANES':
        return this.resolveFreezePanes(activeSheet);
      case 'VISIBILITY_TOGGLE':
        return this.resolveVisibilityToggle(message, captures, activeSheet);
      case 'ROW_COL_STRUCTURE':
        return this.resolveRowColStructure(message, captures, activeSheet);
      default:
        return null;
    }
  }

  private resolveCellFormat(
    message: string,
    captures: RegExpMatchArray,
    activeSheet: string,
  ): Tier0Result | null {
    const formatToken = captures[1]?.toLowerCase();
    const range = parseCellReference(message);
    if (!range || !formatToken) {
      return null;
    }

    const format: FormatSpec = {};
    if (formatToken === 'bold') format.bold = true;
    if (formatToken === 'italic') format.italic = true;
    if (formatToken === 'underline') format.underline = true;

    return {
      skippedLLM: true,
      actions: [
        {
          type: 'FORMAT_RANGE',
          sheetName: activeSheet,
          row: range.row,
          col: range.col,
          rowCount: range.rowCount,
          colCount: range.colCount,
          format,
        },
      ],
    };
  }

  private resolveFreezePanes(activeSheet: string): Tier0Result {
    return {
      skippedLLM: true,
      actions: [{ type: 'FREEZE_PANES', sheetName: activeSheet, freezeRows: 1, freezeColumns: 0 }],
    };
  }

  private resolveVisibilityToggle(
    message: string,
    captures: RegExpMatchArray,
    activeSheet: string,
  ): Tier0Result | null {
    const verb = captures[1]?.toLowerCase();
    const target = captures[2]?.toLowerCase();
    if (!verb || !target) {
      return null;
    }

    if (target === 'sheet') {
      const sheetName = parseSheetName(message) ?? activeSheet;
      if (verb === 'hide') {
        return { skippedLLM: true, actions: [{ type: 'HIDE_SHEET', sheetName }] };
      }
      return { skippedLLM: true, actions: [{ type: 'SHOW_SHEET', sheetName }] };
    }

    if (target === 'row') {
      const rowRef = parseRowReference(message);
      if (!rowRef) {
        return null;
      }
      const type = verb === 'hide' ? 'HIDE_ROW' : 'UNHIDE_ROW';
      return {
        skippedLLM: true,
        actions: [
          {
            type,
            sheetName: activeSheet,
            row: rowRef.row,
            rowCount: rowRef.rowCount,
          },
        ],
      };
    }

    const colRef = parseColumnReference(message);
    if (!colRef) {
      return null;
    }
    const type = verb === 'hide' ? 'HIDE_COLUMN' : 'UNHIDE_COLUMN';
    return {
      skippedLLM: true,
      actions: [
        {
          type,
          sheetName: activeSheet,
          col: colRef.col,
          colCount: colRef.colCount,
        },
      ],
    };
  }

  private resolveRowColStructure(
    message: string,
    captures: RegExpMatchArray,
    activeSheet: string,
  ): Tier0Result | null {
    const verb = captures[1]?.toLowerCase();
    const target = captures[3]?.toLowerCase();
    if (!verb || !target) {
      return null;
    }

    if (target === 'row') {
      const rowRef = parseRowReference(message);
      if (!rowRef) {
        return null;
      }
      const type = verb === 'insert' ? 'INSERT_ROW' : 'DELETE_ROW';
      return {
        skippedLLM: true,
        actions: [
          {
            type,
            sheetName: activeSheet,
            row: rowRef.row,
            rowCount: rowRef.rowCount,
          },
        ],
      };
    }

    const colRef = parseColumnReference(message);
    if (!colRef) {
      return null;
    }

    const type = verb === 'insert' ? 'INSERT_COLUMN' : 'DELETE_COLUMN';
    return {
      skippedLLM: true,
      actions: [
        {
          type,
          sheetName: activeSheet,
          col: colRef.col,
          colCount: colRef.colCount,
        },
      ],
    };
  }
}
