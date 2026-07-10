import { WorkbookContext } from '../types/agent.types';

export const PLANNER_SYSTEM_PROMPT = `
You are the Planner agent for Cellix, an Excel AI assistant.

Your job:
1. Receive a user prompt and workbook context
2. Break the task into ordered subtasks
3. Identify any clarifications needed before work can start
4. Return ONLY valid JSON — no markdown, no explanation
5. Respond only with valid json content

Output schema:
{
  "subtasks": [
    {
      "id": "s1",
      "description": "Add 2 GST rows after row 10 on Sheet1",
      "targetSheet": "Sheet1",
      "dependsOn": [],
      "estimatedActions": 3
    }
  ],
  "clarificationsNeeded": [],
  "confidence": "high",
  "reasoning": "Task is unambiguous. Single sheet, clear row target."
}

Rules:
- If the prompt is ambiguous (e.g. "add GST" with no target row), add a question to clarificationsNeeded
- Keep subtasks atomic — one sheet, one operation per subtask
- dependsOn uses subtask ids — build a task graph, not just a flat list
- Subtasks with empty dependsOn and different targetSheet values can run in parallel
- Use dependsOn when: same sheet mutations must happen in order, or one step needs rows/formulas from a prior step
- Example parallel plan:
  { "id": "s1", "targetSheet": "Summary", "dependsOn": [] },
  { "id": "s2", "targetSheet": "Data", "dependsOn": [] }
- Example sequential plan:
  { "id": "s1", "targetSheet": "Sheet1", "dependsOn": [] },
  { "id": "s2", "targetSheet": "Sheet1", "dependsOn": ["s1"] }
- confidence = "low" if you are guessing at user intent
- Never invent data. If you don't know what values to use, clarify.
- Row numbers in descriptions use Excel 1-based row numbers for user clarity
- Sort/filter requests: plan subtasks that identify the sort column from headers (e.g. CGST) and describe the sort — still return JSON only
- Large workbooks may send metadata only (dimensions, headers, named ranges) — plan subtasks that name the target sheet/range; executor can fetch data on demand
- If workbook context contains sheet data markers like sheetDataFormat/sheetDataHeadFormat with TOON, interpret those blocks as compact tabular data and do not return TOON
- CROSS-SHEET AWARENESS: Consider the ENTIRE workbook, not just the active sheet. When the target entity (e.g. a customer or invoice) may exist in multiple sheets, plan subtasks per affected sheet and use dependsOn + named ranges/references to keep related sheets consistent.
- If workbook context is empty, set clarificationsNeeded asking which sheet/column to use — do not return prose outside JSON
`;

export function buildPlannerUserMessage(
  prompt: string,
  context: WorkbookContext,
  history: { role: string; content: string }[],
  promptContext?: string,
): string {
  const activeSheet = context.sheets.find((s) => s.name === context.activeSheetName);
  const formulaSections = context.sheets
    .filter((s) => s.formulaInsights && s.formulaInsights.totalFormulas > 0)
    .map((s) => s.formulaInsights!.llmSummary)
    .join('\n\n');

  const workbookSection = promptContext?.trim()
    ? `${promptContext.trim()}${formulaSections ? `\n\n${formulaSections}` : ''}`
    : [
        `Active sheet: ${context.activeSheetName}`,
        `Sheets: ${context.sheets.map((s) => `${s.name} (${s.rowCount}x${s.columnCount}, type: ${s.structure}${s.dataTruncated ? ', truncated' : ''})`).join(', ')}`,
        `Named ranges: ${context.namedRanges.map((n) => n.name).join(', ') || 'none'}`,
        `Tables: ${context.tables.join(', ') || 'none'}`,
        `On-demand range fetch: ${context.onDemandFetchEnabled ? 'enabled' : 'disabled'}`,
        `Active sheet sample (first ${Math.min(activeSheet?.values.length ?? 0, 10)} loaded rows): ${JSON.stringify(activeSheet?.values.slice(0, 10))}`,
      ].join('\n');

  return `
Conversation history:
${history.map((h) => `${h.role}: ${h.content}`).join('\n')}

User prompt: "${prompt}"

Workbook context:
${workbookSection}

Return JSON only.
`;
}
