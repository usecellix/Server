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
    HEADER_FORMAT: `
Action hint: HEADER_FORMAT (whole header row — not conditional data rows)
Allowed action types: FORMAT_RANGE only
FORMAT_RANGE schema: { "type": "FORMAT_RANGE", "sheetName": "Purchase Register", "row": 0, "col": 0, "rowCount": 1, "colCount": <number of header columns>, "format": { "fillColor": "#C6EFCE", "bold": true } }
Rules:
- Target EXACTLY the header row: row 0, col 0, rowCount 1, colCount = width of headers / used range
- Light green → "#C6EFCE"; light red → "#FFC7CE"; light yellow → "#FFF2CC"
- NEVER use FORMAT_MATCHING_ROWS (that requires a data-row filter like Payment Status = Pending)
- answer must be preview tense: "I'll highlight the header row…" — never "I've applied"
`,
    CONDITIONAL_FORMAT: `
Action hint: CONDITIONAL_FORMAT
Allowed action types: CONDITIONAL_FORMAT (rule.kind "cellValue" for a numeric column comparison, or "formula" for a cross-column comparison) or FORMAT_MATCHING_ROWS (text/status row match). If the user only wants the header row / headers / first row formatted (no "where/status equals" condition), use FORMAT_RANGE on row 0 instead — never FORMAT_MATCHING_ROWS for header-only requests.

MODIFYING AN EXISTING RULE (critical — check "Existing conditional-format rules" in the user message before every CONDITIONAL_FORMAT action): if the request is changing a rule that already exists on the target sheet/range (e.g. "change the threshold to 15%", "make it top 10 instead of top 5") and one of the listed rules' summary/range matches, add "existingRuleId": "<id from the list>" to the action and set rule to the FULL corrected rule of the SAME kind as the existing one. Do NOT change rule.kind when modifying — if the user wants a different kind of rule entirely, omit existingRuleId and create a new one instead. Never emit a second CONDITIONAL_FORMAT on the same range when an existing rule already covers it and the request is clearly an edit, not an addition.

CONDITIONAL_FORMAT (cellValue) — use whenever the condition is a NUMERIC comparison against a SINGLE column ("above X", "below X", "at least X", "exceeds X" — e.g. "highlight expenses above 1000"). This creates a LIVE Excel rule that re-evaluates automatically when the data changes — never invented as a one-shot fill:
{ "type": "CONDITIONAL_FORMAT", "sheetName": "Purchase Register", "range": "J2:J51", "rule": { "kind": "cellValue", "operator": "greaterThan", "value": 1000, "format": { "fillColor": "#FFC7CE" } } }
Rules:
- range MUST be only the data cells of the single numeric column being compared (exclude the header row, e.g. "J2:J51" not "A1:L51") — never the whole row/table
- rule.operator: greaterThan | greaterThanOrEqual | lessThan | lessThanOrEqual | equalTo | notEqualTo | between | notBetween
- rule.value2 is required only for between/notBetween
- Use this instead of FORMAT_MATCHING_ROWS whenever the comparison value is a number, not a text/status match

CONDITIONAL_FORMAT (formula) — use whenever the condition compares TWO OR MORE columns, or needs a computed threshold that a single column value can't express — e.g. "highlight the regions where revenue dropped more than 10%" (compares this period's revenue against last period's, per row):
{ "type": "CONDITIONAL_FORMAT", "sheetName": "Regional Revenue", "range": "A2:D9", "rule": { "kind": "formula", "formula": "=$C2<$B2*0.9", "format": { "fillColor": "#FFC7CE" } } }
Rules:
- formula is a boolean Excel formula evaluated relative to the TOP-LEFT cell of range — e.g. if range starts at row 2, write the formula in terms of row 2 (Excel applies it to every row in range with relative references shifting automatically)
- $-anchor the COLUMN of any cell reference you want fixed while the row varies down the range (e.g. "$B2", "$C2") — this is what makes one formula apply correctly to every row
- range should cover every column the formula reads plus every cell that should be highlighted when the row matches (usually the full row span of the data, not just one column)
- Never invent a numeric column threshold here — if the comparison only involves one column against a constant, use the cellValue variant above instead

CONDITIONAL_FORMAT (topBottom) — use whenever the request is rank-based rather than a fixed threshold ("highlight the top 5 suppliers by total", "flag the bottom 10% of scores"):
{ "type": "CONDITIONAL_FORMAT", "sheetName": "Suppliers", "range": "C2:C40", "rule": { "kind": "topBottom", "side": "top", "rank": 5, "format": { "fillColor": "#C6EFCE" } } }
Rules:
- range MUST be only the data cells of the single numeric column being ranked (exclude the header row)
- rule.side: "top" | "bottom"; rule.rank: item count by default, or a 0-100 percentage when rule.isPercent is true
- Never use cellValue/formula here — "top N" / "bottom N%" is not a fixed threshold, it re-ranks as data changes

CONDITIONAL_FORMAT (colorScale) — use for a gradient/heat-map request ("color-scale the Total Amount column", "add a heat map to the scores"), never a discrete threshold or rank:
{ "type": "CONDITIONAL_FORMAT", "sheetName": "Scores", "range": "B2:B50", "rule": { "kind": "colorScale", "colors": ["#F8696B", "#FFEB84", "#63BE7B"] } }
Rules:
- range MUST be only the data cells of the single numeric column being scaled (exclude the header row)
- rule.colors: 2 hex colors (low→high) or 3 hex colors (low→mid→high) — no "format" field, colorScale has no bold/fill toggle, only the gradient
- The scale's low/high ends are the range's current actual min/max and shift automatically as data changes

FORMAT_MATCHING_ROWS schema — use this for TEXT/STATUS row matches only (e.g. "Payment Status is Pending"), never for numeric comparisons (Office.js reads the sheet and paints matching rows — never invent per-row HIGHLIGHT_CELL lists or a "condition" field on HIGHLIGHT_CELL/FORMAT_RANGE):
Apply highlight: { "type": "FORMAT_MATCHING_ROWS", "sheetName": "Purchase Register", "range": "A1:L51", "hasHeaders": true, "filter": { "column": "Payment Status", "operator": "equals", "value": "Pending" }, "format": { "fillColor": "#FFC7CE" } }
Remove/clear highlight: { "type": "FORMAT_MATCHING_ROWS", "sheetName": "Purchase Register", "range": "A1:L51", "hasHeaders": true, "filter": { "column": "Payment Status", "operator": "equals", "value": "Pending" }, "format": { "clearFill": true } }
Rules:
- filter.column MUST be the header name (e.g. "Payment Status"), never a numeric index
- filter.operator: equals | contains | notEquals | greaterThan | lessThan — but if value is numeric AND operator is greaterThan/lessThan, use CONDITIONAL_FORMAT above instead
- Use the sheet's used range for "range" (include the header row)
- To REMOVE highlights/fill: use format.clearFill true — NEVER white fill (#FFFFFF) and NEVER enumerate per-row FORMAT_RANGE actions
- Light red → "#FFC7CE"; light yellow → "#FFF2CC"; light green → "#C6EFCE"
- Emit exactly ONE action (CONDITIONAL_FORMAT or FORMAT_MATCHING_ROWS) for a highlight request (hasHeaders: true always on FORMAT_MATCHING_ROWS)
- answer must be preview tense: "I'll highlight…" — never "I've applied"
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
  "answer": "one short sentence in preview tense (I'll …) for the Accept/Reject UI — never past tense I've applied",
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

  const existingRules = (workbookContext.conditionalFormats ?? []).filter(
    (rule) => rule.sheetName === workbookContext.activeSheetName,
  );
  const existingRulesLine =
    actionHint === 'CONDITIONAL_FORMAT'
      ? `Existing conditional-format rules on this sheet: ${
          existingRules.length > 0
            ? existingRules
                .map((rule) => `[${rule.id}] ${rule.range} (${rule.ruleKind}: ${rule.summary})`)
                .join('; ')
            : '(none)'
        }\n`
      : '';

  return [
    `Action hint: ${actionHint}`,
    `Active sheet: ${workbookContext.activeSheetName}`,
    `Used range: ${usedRange}`,
    `Headers: ${headers.length > 0 ? headers.join(', ') : '(unknown)'}`,
    `Sample rows (first 5 data rows): ${JSON.stringify(sampleRows.slice(1))}`,
    existingRulesLine,
    `User message: "${message}"`,
  ]
    .filter((line) => line !== '')
    .join('\n');
}
