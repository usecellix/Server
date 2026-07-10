import { Action, SubTask, WorkbookContext } from '../types/agent.types';

export const EXECUTOR_SYSTEM_PROMPT = `
You are the Executor agent for Cellix, an Excel AI assistant.

Your job:
- Receive ONE subtask and workbook context (may be metadata-first / compressed for large sheets)
- Emit the exact typed actions needed to complete it
- Request additional cell data via get_range_data when compressed context is insufficient
- Return ONLY valid JSON — no markdown, no explanation
- Respond only with valid json content

Available action types (0-based row/col in JSON; row 0 = Excel row 1 = header):
SET_CELL, SET_FORMULA, ADD_ROW, DELETE_ROW, INSERT_ROW, INSERT_COLUMN, DELETE_COLUMN,
FORMAT_RANGE, FILL_DOWN, FILL_RIGHT, WRITE_TABLE, CREATE_SHEET, DELETE_SHEET, RENAME_SHEET, COPY_SHEET,
BATCH_SET, CREATE_TABLE, DEFINE_NAMED_RANGE, AUTOFIT_COLUMNS, ADD_SHEET,
MERGE_CELLS, CLEAR_CONTENT, HIGHLIGHT_CELL, SORT_RANGE

On-demand data tool — use when sheet data is truncated or you need rows not in context:
{
  "subtaskId": "s1",
  "actions": [],
  "isDone": false,
  "toolRequest": { "name": "get_range_data", "sheet": "Sheet1", "range": "A1:H200" }
}
- Request only the range you need (include headers when sorting/filtering)
- After fetch, the next turn will include the fetched values — then emit actions

SORT_RANGE schema (for sort/filter requests):
{ "type": "SORT_RANGE", "sheetName": "Sheet1", "range": "A1:H50", "key": 3, "ascending": true, "hasHeaders": true }
- key = 0-based column index within range (header row identifies columns)
- hasHeaders: true when first row is headers (do not sort header into data)

Output schema:
{
  "subtaskId": "s1",
  "actions": [
    { "type": "ADD_ROW", "data": ["GST", "", "=C10*0.1"] }
  ],
  "isDone": true,
  "nextStep": null
}

Rules:
- Use 0-based row/col indices in JSON (row 0 = header row in Excel row 1)
- ADD_ROW appends a new data row with a data array aligned to columns
- Formulas must be valid Excel syntax (include leading =)
- If the subtask is too vague to act on, return isDone: false and describe what's missing in nextStep
- Never delete data unless the subtask explicitly says to
- For SET_FORMULA referencing cells outside the current row, use absolute references ($A$1)
- Set sheetName on actions when targeting a non-active sheet
- For large sheets: check dimensions vs visible rows — use toolRequest before SORT_RANGE or row-specific edits
- If context includes sheetDataFormat/sheetDataHeadFormat as TOON, interpret it as compact tabular data and never return TOON in output
`;

function formatSparseSheetPreview(sheet: WorkbookContext['sheets'][number]): string {
  const lines = [
    `Sheet "${sheet.name}": ${sheet.rowCount}x${sheet.columnCount}, range ${sheet.usedRange}, structure ${sheet.structure}`,
  ];

  if (sheet.dataTruncated || sheet.compressionMeta?.truncated) {
    lines.push(
      `DATA TRUNCATED — only header + sample rows loaded. Full sheet has ${sheet.rowCount} rows. Use toolRequest to fetch needed ranges.`,
    );
  }

  const previewRows = Math.min(sheet.values.length, 15);
  lines.push(`Visible values (first ${previewRows} loaded rows): ${JSON.stringify(sheet.values.slice(0, previewRows))}`);

  if (sheet.rowCount > sheet.values.length) {
    lines.push(`... ${sheet.rowCount - sheet.values.length} rows not loaded — fetch with get_range_data`);
  }

  return lines.join('\n');
}

export function buildExecutorUserMessage(
  subtask: SubTask,
  context: WorkbookContext,
  previousActions: Action[],
): string {
  const normalizedTarget = subtask.targetSheet.replace(/@\[(.+?)\]/g, '$1').trim();
  const targetSheet =
    context.sheets.find((s) => s.name === normalizedTarget) ??
    context.sheets.find((s) => s.name.toLowerCase() === normalizedTarget.toLowerCase());
  const feedbackBlock = [
    context.verifierFeedback
      ? `Verifier feedback from previous attempt: ${context.verifierFeedback}\nIssues: ${JSON.stringify(context.verifierIssues ?? [])}`
      : '',
    context.formulaValidationFeedback
      ? `Formula validator feedback: ${context.formulaValidationFeedback}\nFormula issues: ${JSON.stringify(context.formulaValidationIssues ?? [])}`
      : '',
  ]
    .filter(Boolean)
    .join('\n');

  const formulaBlock = targetSheet?.formulaInsights?.llmSummary
    ? `\nFormula analysis:\n${targetSheet.formulaInsights.llmSummary}\n`
    : '';

  const fetchedBlock =
    context.fetchedRanges && context.fetchedRanges.length > 0
      ? `\nRanges fetched this session:\n${context.fetchedRanges.map((r) => `- ${r.sheet}!${r.range} (${r.rowCount} rows)`).join('\n')}\n`
      : '';

  const sheetBlock = targetSheet ? formatSparseSheetPreview(targetSheet) : 'Target sheet not found in context';

  return `
${feedbackBlock ? `${feedbackBlock}\n` : ''}
Subtask: ${subtask.description}
Target sheet: ${subtask.targetSheet}
On-demand fetch available: ${context.onDemandFetchEnabled ? 'yes' : 'no'}
${fetchedBlock}
${formulaBlock}
${sheetBlock}

Sheet formulas (loaded rows):
${JSON.stringify(targetSheet?.formulas.slice(0, Math.min(targetSheet?.formulas.length ?? 0, 15)))}

Number formats (first row):
${JSON.stringify(targetSheet?.numberFormats[0])}

Actions already applied in this session:
${JSON.stringify(previousActions)}

Return JSON only.
`;
}
