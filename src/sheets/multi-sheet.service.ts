import { Injectable, Logger } from '@nestjs/common';
import { SheetSnapshot, WorkbookContext } from '../types/cellix.types';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import { SheetActionPayload } from '../excel-ai/types/sheet-actions.types';

export interface CellDiff {
  address: string;
  valueA: string;
  valueB: string;
}

export interface CompareResult {
  sheetA: string;
  sheetB: string;
  summary: string;
  addedInB: string[];
  removedInB: string[];
  modifiedCells: CellDiff[];
  totalDiffs: number;
  differences: [];
}

const SHEET_CREATE_TYPES = new Set<SheetActionPayload['type']>([
  'CREATE_SHEET',
  'ADD_SHEET',
  'COPY_SHEET',
]);

function createdSheetName(action: SheetActionPayload): string | null {
  if (action.type === 'CREATE_SHEET' || action.type === 'ADD_SHEET') {
    const name = action.name ?? action.sheetName;
    return typeof name === 'string' && name.trim() ? name.trim() : null;
  }
  if (action.type === 'COPY_SHEET' || action.type === 'RENAME_SHEET') {
    const name = action.newName ?? action.newSheetName;
    return typeof name === 'string' && name.trim() ? name.trim() : null;
  }
  return null;
}

/**
 * Validate that cross-sheet actions reference sheets present in the workbook context
 * or created earlier in the same action batch.
 */
export function validateCrossSheetActions(
  actions: SheetActionPayload[],
  context: WorkbookContext,
): { valid: SheetActionPayload[]; invalid: SheetActionPayload[]; errors: string[] } {
  const knownSheets = new Set(context.sheets.map((sheet) => sheet.sheetName));

  // Sheets created in this batch are valid targets for later actions (ADD_ROW, SET_CELL, …).
  for (const action of actions) {
    const created = createdSheetName(action);
    if (created) knownSheets.add(created);
  }

  const valid: SheetActionPayload[] = [];
  const invalid: SheetActionPayload[] = [];
  const errors: string[] = [];

  for (const action of actions) {
    if (SHEET_CREATE_TYPES.has(action.type) || action.type === 'RENAME_SHEET') {
      valid.push(action);
      continue;
    }

    const sheet = action.sheetName ?? context.activeSheet;
    if (sheet && !knownSheets.has(sheet)) {
      invalid.push(action);
      errors.push(`Action "${action.type}" references unknown sheet "${sheet}"`);
    } else {
      valid.push(action);
    }
  }

  return { valid, invalid, errors };
}

function colIndexToLetter(col: number): string {
  let index = col + 1;
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function buildFallbackSummary(
  addedInB: string[],
  removedInB: string[],
  modified: CellDiff[],
): string {
  const parts: string[] = [];
  if (addedInB.length > 0) parts.push(`${addedInB.length} row(s) added`);
  if (removedInB.length > 0) parts.push(`${removedInB.length} row(s) removed`);
  if (modified.length > 0) parts.push(`${modified.length} cell(s) modified`);
  return parts.length > 0 ? `${parts.join(', ')}.` : 'No differences found.';
}

@Injectable()
export class MultiSheetService {
  private readonly logger = new Logger(MultiSheetService.name);

  constructor(private readonly openRouterService: OpenRouterService) {}

  async compareSheets(
    snapshotA: SheetSnapshot,
    snapshotB: SheetSnapshot,
  ): Promise<CompareResult> {
    const modifiedCells: CellDiff[] = [];
    const addedInB: string[] = [];
    const removedInB: string[] = [];

    const rowsA = snapshotA.sampleData;
    const rowsB = snapshotB.sampleData;
    const colCount = Math.max(snapshotA.colCount, snapshotB.colCount);

    const keysA = new Set(rowsA.map((row) => String(row[0] ?? '')));
    const keysB = new Set(rowsB.map((row) => String(row[0] ?? '')));

    for (const row of rowsB) {
      const key = String(row[0] ?? '');
      if (key && !keysA.has(key)) {
        addedInB.push(key);
      }
    }

    for (const row of rowsA) {
      const key = String(row[0] ?? '');
      if (key && !keysB.has(key)) {
        removedInB.push(key);
      }
    }

    for (let rowIndex = 0; rowIndex < Math.min(rowsA.length, rowsB.length); rowIndex += 1) {
      const rowA = rowsA[rowIndex];
      const rowB = rowsB[rowIndex];
      const keyA = String(rowA[0] ?? '');
      const keyB = String(rowB[0] ?? '');

      if (keyA !== keyB) continue;

      for (let colIndex = 0; colIndex < colCount; colIndex += 1) {
        const valueA = normalizeCell(rowA[colIndex]);
        const valueB = normalizeCell(rowB[colIndex]);
        if (valueA !== valueB) {
          const rowNum = rowIndex + 2;
          modifiedCells.push({
            address: `${colIndexToLetter(colIndex)}${rowNum}`,
            valueA: valueA || '(empty)',
            valueB: valueB || '(empty)',
          });
        }
      }
    }

    const summary = await this.generateCompareSummary(
      snapshotA,
      snapshotB,
      addedInB,
      removedInB,
      modifiedCells,
    );

    return {
      sheetA: snapshotA.sheetName,
      sheetB: snapshotB.sheetName,
      summary,
      addedInB,
      removedInB,
      modifiedCells,
      totalDiffs: addedInB.length + removedInB.length + modifiedCells.length,
      differences: [],
    };
  }

  private async generateCompareSummary(
    snapshotA: SheetSnapshot,
    snapshotB: SheetSnapshot,
    addedInB: string[],
    removedInB: string[],
    modified: CellDiff[],
  ): Promise<string> {
    const prompt = `Compare two Excel sheets and summarize the differences in one concise sentence.

Sheet A: "${snapshotA.sheetName}" — ${snapshotA.rowCount} rows
Sheet B: "${snapshotB.sheetName}" — ${snapshotB.rowCount} rows

Differences:
- Added in B: ${addedInB.length} row(s)${addedInB.length > 0 ? ` (${addedInB.slice(0, 3).join(', ')})` : ''}
- Removed in B: ${removedInB.length} row(s)${removedInB.length > 0 ? ` (${removedInB.slice(0, 3).join(', ')})` : ''}
- Modified cells: ${modified.length}

Write one short summary sentence (max 20 words). No markdown.`;

    try {
      const summary = await this.openRouterService.quickCall(
        'You summarize Excel sheet differences in one sentence.',
        prompt,
      );
      if (summary.trim()) {
        return summary.trim();
      }
    } catch (error: unknown) {
      const detail = error instanceof Error ? error.message : 'compare summary failed';
      this.logger.warn(`Compare summary LLM failed: ${detail}`);
    }

    return buildFallbackSummary(addedInB, removedInB, modified);
  }
}
