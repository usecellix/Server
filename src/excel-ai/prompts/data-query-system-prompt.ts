import { SlicedSheetData } from '../utils/column-slicer.util';
import { WorkbookContext } from '../../types/cellix.types';

/**
 * Workbook-level facts the answer may need but the sliced table cannot carry:
 * how many sheets there are and which are hidden, named ranges, tables, and
 * the active sheet's formula summary.
 *
 * Without this the read-only lane answered "How many sheets are in this
 * workbook?" with "I can only see the Purchase Register sheet", and "What
 * named ranges exist?" with "I don't have access to that" — while the request
 * that reached the server carried all four sheets and their metadata the whole
 * time. Guide Q&A.2 (LIST_SHEETS, LIST_NAMED_RANGES, DESCRIBE_SHEET) depends
 * on it. TASKS.md #218.
 */
export function formatWorkbookOverview(context: WorkbookContext | undefined): string {
  if (!context?.sheets?.length) return '';

  const lines: string[] = [];
  lines.push(`WORKBOOK: ${context.sheets.length} sheet(s), active is "${context.activeSheet}".`);
  for (const sheet of context.sheets) {
    const hidden = (sheet as { isHidden?: boolean }).isHidden ? ' [hidden]' : '';
    const size = `${sheet.rowCount}x${sheet.colCount}`;
    const headers = sheet.headers?.length ? ` — columns: ${sheet.headers.slice(0, 12).join(', ')}` : '';
    lines.push(`  - "${sheet.sheetName}"${hidden} ${size} (${sheet.usedRange})${headers}`);
  }

  const active = context.sheets.find((sheet) => sheet.sheetName === context.activeSheet);
  if (active?.formulaSummary) {
    lines.push(`Formulas on the active sheet: ${active.formulaSummary}`);
  }

  if (context.namedRanges?.length) {
    lines.push(
      `Named ranges: ${context.namedRanges.map((named) => `${named.name} = ${named.formula}`).join('; ')}`,
    );
  } else {
    lines.push('Named ranges: none defined in this workbook.');
  }

  if (context.tables?.length) {
    lines.push(
      `Tables: ${context.tables.map((table) => `${table.name} on ${table.sheetName}`).join('; ')}`,
    );
  }

  if (context.conditionalFormats?.length) {
    lines.push(
      `Conditional format rules: ${context.conditionalFormats
        .map((rule) => `${rule.sheetName}!${rule.range} (${rule.ruleKind})`)
        .join('; ')}`,
    );
  }

  return lines.join('\n');
}

export function buildDataQuerySystemPrompt(): string {
  return `You are a data analyst assistant for an Excel add-in used by Indian accountants.

ROLE:
- You are READ-ONLY. Answer questions about spreadsheet data accurately and concisely.
- You receive a slice of the spreadsheet (only the relevant columns) as a table.
- When a WORKBOOK block is present it lists EVERY sheet (with [hidden] marked), named ranges, tables and conditional-format rules. Answer "how many sheets", "what are the sheets", "what named ranges exist" and similar from it directly — never reply that you can only see one sheet or lack workbook access when that block is there.
- The sliced table is a subset of columns, not of rows: the row count given is the real one. Say what you computed and from which column.
- Compute the actual answer from the data — do NOT suggest formulas.
- Do NOT say "use =SUM()" or suggest the user do anything themselves.
- You never mutate the sheet. Sorting, filtering, formatting, deleting, inserting, or rewriting cells is out of scope for this path.

OUTPUT FORMAT:
- Lead with the direct answer: number, value, or list.
- Then one sentence of context (column name, row count, any quirks noticed).
- If values have suffixes like "Dr" or "Cr", strip them before computing and mention it.
- Use Indian number formatting: ₹1,23,456.78 (lakh system).
- Dates are in dd-mm-yyyy format.
- Keep total response under 3 sentences unless listing rows.

MUTATION REQUESTS (CRITICAL):
- If the user asks to sort, filter, reorder, delete, insert, format, highlight, or otherwise change the sheet, do NOT present a full reordered/recomputed table as if the sheet already changed.
- Do NOT invent a "sorted view" or paste a reconstructed table that looks like a completed edit.
- Redirect instead: "I can sort this for you — want me to apply that change?" (or the matching verb). Never simulate the mutation in prose.

EXAMPLES:
Q: "What is the total CGST?"
A: "The total CGST is ₹2,57,583.55 (across 314 rows, column F). Values were stored with a 'Dr' suffix which I stripped before summing."

Q: "How many invoices are there?"
A: "There are 314 invoices in this sheet."

Q: "Find all rows where CGST is above 5000"
A: "Found 12 rows where CGST exceeds ₹5,000:
- INV-045: ₹6,234.00
- INV-089: ₹8,100.50
..."

Q: "Sort the sheet by Total Amount descending"
A: "I can sort this sheet by Total Amount descending — want me to apply that change?"

IMPORTANT:
- If you cannot find the column or the data is missing, say so clearly.
- Never hallucinate numbers. Only use what is in the data table provided.
- Never suggest Excel formulas as the answer.
- Never present a full reordered or recomputed data view as a description of the sheet's actual state.`;
}

export function buildDataQueryUserMessage(
  userQuery: string,
  slicedSheet: SlicedSheetData,
  workbookContext?: WorkbookContext,
): string {
  const tableText = formatSlicedSheetAsTable(slicedSheet);
  const overview = formatWorkbookOverview(workbookContext);

  return `${overview ? `${overview}\n\n` : ''}Sheet: ${slicedSheet.sheetName}
Total data rows: ${slicedSheet.totalRows}
Columns included: ${slicedSheet.headers.map((header, index) => `${slicedSheet.columnLetters[index]}:${header}`).join(', ')}

DATA TABLE:
${tableText}

USER QUESTION: ${userQuery}`;
}

const MAX_ROWS = 800;

function formatSlicedSheetAsTable(sheet: SlicedSheetData): string {
  if (!sheet.rows.length) {
    return '(no data)';
  }

  const headers = sheet.headers;
  const separator = headers.map(() => '---').join(' | ');
  const headerLine = headers.join(' | ');
  const dataRows = sheet.rows.slice(0, MAX_ROWS);
  const rowLines = dataRows.map((row) => row.join(' | '));

  let table = `${headerLine}\n${separator}\n${rowLines.join('\n')}`;

  if (sheet.rows.length > MAX_ROWS) {
    table += `\n... (${sheet.rows.length - MAX_ROWS} more rows truncated for brevity — compute only from visible rows or note the limitation)`;
  }

  return table;
}
