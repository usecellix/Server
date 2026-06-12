import { ConversationTurn, WorkbookContext } from '../../types/cellix.types';
import { extractFormatContext } from '../excel/format-context-reader';

export const FORMAT_INHERITANCE_RULES = `## Format inheritance — CRITICAL
- For date columns, ALWAYS include "format": {"numberFormat": "<detected format>"} on SET_CELL, SET_FORMULA, ADD_ROW, and INSERT_ROW actions
- New rows inherit format from the row above automatically on the client — still include explicit numberFormat when writing date or currency values
- When adding rows (ADD_ROW / INSERT_ROW), include format.numberFormat for any column that has a detected date or currency format
- Do NOT change number formats on columns unless the user explicitly asked to reformat`;

export const ACTION_TYPES_REFERENCE = `
SET_CELL:     { type, row, col, value, formula?, format? }
SET_FORMULA:  { type, row, col, formula, format? }
ADD_ROW:      { type, data[], format? }
INSERT_ROW:   { type, row, count?, position?, data?, format? }
DELETE_ROW:   { type, row }
INSERT_COLUMN:{ type, col, count?, position? }
DELETE_COLUMN:{ type, col }
FORMAT_RANGE: { type, row, col, rowCount?, colCount?, format }
WRITE_TABLE:  { type, headers[], rows[][] }
CREATE_SHEET: { type, sheetName }
RENAME_SHEET: { type, newSheetName }
COPY_SHEET:   { type, sheetName, newSheetName }
MERGE_CELLS:  { type, row, col, rowCount, colCount }
UNMERGE_CELLS:{ type, row, col, rowCount, colCount }
CLEAR_CONTENT:{ type, row, col, rowCount?, colCount? }
CLEAR_FORMAT: { type, row, col, rowCount?, colCount? }
FILL_DOWN:    { type, col, row, endRow }
FILL_RIGHT:   { type, row, col, endCol }
`.trim();

function formatSampleRow(row: (string | number | null)[]): string {
  return row.map((cell) => (cell == null ? '' : String(cell))).join(' | ');
}

export function buildMultiSheetNotes(context: WorkbookContext): string {
  if (context.sheets.length <= 1) {
    return '';
  }

  const sheetList = context.sheets.map((sheet) => `"${sheet.sheetName}"`).join(', ');
  return `
## Multi-sheet notes
You have ${context.sheets.length} sheets in context: ${sheetList}.
The active sheet is "${context.activeSheet}".
If the user's request affects a specific sheet, set the "sheetName" field on each action to the correct sheet name.
If the user's request is cross-sheet (e.g. "copy column A from Sheet1 to Sheet2"), generate actions for BOTH sheets.`.trim();
}

/**
 * Builds the format-aware workbook section appended to the main Cellix system prompt.
 * Deterministic: columns are sorted by index; sheet order follows context.sheets.
 */
export function buildFormatContextSection(context: WorkbookContext): string {
  const formatContexts = extractFormatContext(context);

  const sheetDescriptions = context.sheets
    .map((sheet) => {
      const fmt = formatContexts.find((f) => f.sheetName === sheet.sheetName);
      if (!fmt) return '';

      const headerRow = sheet.headers.join(' | ');
      const sampleRows = sheet.sampleData
        .slice(0, 5)
        .map((row) => formatSampleRow(row))
        .join('\n');
      const formatRules = fmt.columns
        .map((col) => `  - Col ${col.columnIndex + 1} "${col.header}": ${col.rule}`)
        .join('\n');
      const lastDataAddr = `Row ${fmt.lastDataRow}`;
      const isActive = sheet.sheetName === context.activeSheet ? 'YES' : 'no';

      return `
### Sheet: "${sheet.sheetName}" (${sheet.rowCount} rows × ${sheet.colCount} cols)
Headers (Row 1): ${headerRow}
Sample data:
${sampleRows}
Format rules (MUST follow exactly):
${formatRules}
Last data row: ${lastDataAddr}
Active sheet: ${isActive}`.trim();
    })
    .filter(Boolean)
    .join('\n\n---\n\n');

  if (!sheetDescriptions) {
    return '';
  }

  const multiSheetNotes = buildMultiSheetNotes(context);

  return `## Workbook format context
${sheetDescriptions}

${FORMAT_INHERITANCE_RULES}${multiSheetNotes ? `\n\n${multiSheetNotes}` : ''}`;
}

export function buildSystemPrompt(
  context: WorkbookContext,
  conversationHistory: ConversationTurn[] = [],
): string {
  const formatSection = buildFormatContextSection(context);
  const historyBlock =
    conversationHistory.length > 0
      ? `\nConversation history:\n${conversationHistory.map((t) => `${t.role}: ${t.content}`).join('\n')}\n`
      : '';

  return `You are Cellix, an AI assistant that manipulates Excel workbooks.

## Your job
Analyze the user's request, understand the workbook structure, and return a JSON object describing exactly what changes to make.

${formatSection}
${historyBlock}

## Output rules — CRITICAL
1. Respond ONLY with a valid JSON object: {"type":"actions","answer":"...","explanation":"...","actions":[...]}
2. Do NOT include markdown fences, explanations, or any text outside the JSON
3. For date columns, ALWAYS include "format": {"numberFormat": "<detected format>"} in write actions
4. Use 0-based row/col indices (row 0 = Excel row 1 = header row)

## Action types available
${ACTION_TYPES_REFERENCE}
`;
}
