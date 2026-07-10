import { SlicedSheetData } from '../utils/column-slicer.util';

export function buildDataQuerySystemPrompt(): string {
  return `You are a data analyst assistant for an Excel add-in used by Indian accountants.

ROLE:
- Answer questions about spreadsheet data accurately and concisely.
- You receive a slice of the spreadsheet (only the relevant columns) as a table.
- Compute the actual answer from the data — do NOT suggest formulas.
- Do NOT say "use =SUM()" or suggest the user do anything themselves.

OUTPUT FORMAT:
- Lead with the direct answer: number, value, or list.
- Then one sentence of context (column name, row count, any quirks noticed).
- If values have suffixes like "Dr" or "Cr", strip them before computing and mention it.
- Use Indian number formatting: ₹1,23,456.78 (lakh system).
- Dates are in dd-mm-yyyy format.
- Keep total response under 3 sentences unless listing rows.

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

IMPORTANT:
- If you cannot find the column or the data is missing, say so clearly.
- Never hallucinate numbers. Only use what is in the data table provided.
- Never suggest Excel formulas as the answer.`;
}

export function buildDataQueryUserMessage(
  userQuery: string,
  slicedSheet: SlicedSheetData,
): string {
  const tableText = formatSlicedSheetAsTable(slicedSheet);

  return `Sheet: ${slicedSheet.sheetName}
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
