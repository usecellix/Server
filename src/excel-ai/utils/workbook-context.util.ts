import { SheetAnalysis } from '../services/sheet-analyzer.service';
import { WorkbookContext } from '../types/sheet-actions.types';

export function buildWorkbookContext(
  sheetData: unknown[][],
  analysis: SheetAnalysis,
  options?: { activeSheet?: string; sheets?: string[] },
): WorkbookContext {
  // 1-based Excel row number. analysis.headerRowIndex is 0-based and defaults to 0
  // (row 1) for sheets where SheetAnalyzerService couldn't detect a better row —
  // this no longer hardcodes row 1 regardless of what was actually detected.
  const headerRow = (analysis.headerRowIndex ?? 0) + 1;
  const dataStartRow = headerRow + 1;
  const lastDataRow = Math.max(analysis.rowCount, headerRow);
  const totalRows = Math.max(0, lastDataRow - headerRow);

  const headers: Record<string, string> = {};
  analysis.columnLetters.forEach((letter, index) => {
    headers[letter] = analysis.headers[index] || `Column ${letter}`;
  });

  const endCol = analysis.columnLetters[analysis.columnCount - 1] ?? 'A';
  // Range starts at the detected header row, not always A1 — rows above it are
  // preamble/title rows, not part of the table.
  const dataRange = analysis.isEmpty ? 'A1' : `A${headerRow}:${endCol}${lastDataRow}`;

  return {
    activeSheet: options?.activeSheet ?? 'Sheet1',
    sheets: options?.sheets ?? [options?.activeSheet ?? 'Sheet1'],
    headers,
    dataRange,
    lastDataRow,
    headerRow,
    dataStartRow,
    totalRows,
    columnCount: analysis.columnCount,
    columnLetters: analysis.columnLetters,
  };
}

export function formatWorkbookContextForPrompt(ctx: WorkbookContext): string {
  const headerLines = Object.entries(ctx.headers)
    .map(([letter, name]) => `  "${letter}": "${name}"`)
    .join(',\n');

  return `WorkbookContext:
{
  "activeSheet": "${ctx.activeSheet}",
  "sheets": ${JSON.stringify(ctx.sheets)},
  "headers": {
${headerLines}
  },
  "dataRange": "${ctx.dataRange}",
  "lastDataRow": ${ctx.lastDataRow},
  "headerRow": ${ctx.headerRow},
  "dataStartRow": ${ctx.dataStartRow},
  "totalRows": ${ctx.totalRows}
}`;
}
