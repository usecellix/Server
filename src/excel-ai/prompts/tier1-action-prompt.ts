import { WorkbookContext } from '../../agents/types/agent.types';

const SHARED_RULES = `
Respond ONLY with valid JSON — no markdown fences.
Produce exactly ONE action in the "actions" array.
Use 0-based row/col indices. Include sheetName when the action targets a sheet.
Never invent hard-coded numeric totals where a formula is expected.
`;

export function buildTier1SystemPrompt(actionHint: string): string {
  const actionRules: Record<string, string> = {
    SORT_OR_FILTER: `
Action hint: SORT_OR_FILTER
Allowed action types: SORT_RANGE
SORT_RANGE schema: { "type": "SORT_RANGE", "sheetName": "...", "range": "A1:H50", "key": 0, "ascending": true, "hasHeaders": true, "columnName": "..." }
For filter-only requests with no sort, use SORT_RANGE on the header row range as a best-effort visible reorder, or return a CLARIFY action if the target column is unknown.
`,
    FIND_REPLACE: `
Action hint: FIND_REPLACE (text columns only — numeric/financial columns must never reach this path)
Allowed action types: BATCH_SET
BATCH_SET schema: { "type": "BATCH_SET", "sheetName": "...", "operations": [{ "address": "A2", "value": "new text" }] }
Only replace explicit text tokens the user named — do not rewrite formulas or currency columns.
`,
    CONDITIONAL_FORMAT: `
Action hint: CONDITIONAL_FORMAT
Allowed action types: FORMAT_MATCHING_ROWS
FORMAT_MATCHING_ROWS schema (Office.js reads the sheet and paints matching rows — never invent per-row HIGHLIGHT_CELL lists or a "condition" field on HIGHLIGHT_CELL/FORMAT_RANGE):
Apply highlight: { "type": "FORMAT_MATCHING_ROWS", "sheetName": "Purchase Register", "range": "A1:L51", "hasHeaders": true, "filter": { "column": "Payment Status", "operator": "equals", "value": "Pending" }, "format": { "fillColor": "#FFC7CE" } }
Remove/clear highlight: { "type": "FORMAT_MATCHING_ROWS", "sheetName": "Purchase Register", "range": "A1:L51", "hasHeaders": true, "filter": { "column": "Payment Status", "operator": "equals", "value": "Pending" }, "format": { "clearFill": true } }
Rules:
- filter.column MUST be the header name (e.g. "Payment Status"), never a numeric index
- filter.operator: equals | contains | notEquals | greaterThan | lessThan
- Use the sheet's used range for "range" (include the header row)
- To REMOVE highlights/fill: use format.clearFill true — NEVER white fill (#FFFFFF) and NEVER enumerate per-row FORMAT_RANGE actions
- Light red → "#FFC7CE"; light yellow → "#FFF2CC"; light green → "#C6EFCE"
- Emit exactly ONE FORMAT_MATCHING_ROWS action — do not use HIGHLIGHT_CELL or FORMAT_RANGE for row-conditional highlights
`,
    COPY_FILL: `
Action hint: COPY_FILL
Allowed action types: FILL_DOWN, FILL_RIGHT, FORMAT_RANGE
FILL_DOWN schema: { "type": "FILL_DOWN", "sheetName": "...", "col": 0, "row": 1, "endRow": 20 }
For copy formatting, use FORMAT_RANGE with the format copied from the source cell context.
`,
    GENERIC: `
Infer the single best SheetAction for the user request from the workbook context.
`,
  };

  const specific = actionRules[actionHint] ?? actionRules.GENERIC;

  return `You are Cellix Tier-1 — produce exactly one low-stakes Excel write action.
${SHARED_RULES}
${specific}
Response JSON shape:
{
  "answer": "one sentence for the user",
  "actions": [ { "type": "...", ... } ]
}`;
}

export function buildTier1UserMessage(
  message: string,
  actionHint: string,
  workbookContext: WorkbookContext,
): string {
  const activeSheet = workbookContext.sheets.find(
    (sheet) => sheet.name === workbookContext.activeSheetName,
  );
  const headers = (activeSheet?.values[0] ?? []).map((cell) => String(cell ?? '').trim());
  const usedRange = activeSheet?.usedRange ?? 'A1';
  const sampleRows = (activeSheet?.values ?? []).slice(0, 6);

  return [
    `Action hint: ${actionHint}`,
    `Active sheet: ${workbookContext.activeSheetName}`,
    `Used range: ${usedRange}`,
    `Headers: ${headers.length > 0 ? headers.join(', ') : '(unknown)'}`,
    `Sample rows (first 5 data rows): ${JSON.stringify(sampleRows.slice(1))}`,
    `User message: "${message}"`,
  ].join('\n');
}
