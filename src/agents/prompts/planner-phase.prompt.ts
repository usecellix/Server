import { WorkbookContext } from '../types/agent.types';

/**
 * Two-pass planning, pass 1 (TASKS.md #191) — identify the top-level PHASES
 * of a build, not the concrete subtasks. Deliberately tiny: a coarse phase
 * list for even the largest realistic request (12 month sheets + a
 * multi-section dashboard) is 5-8 short entries, nowhere near a token budget
 * that risks truncation. The real detail (formulas, dropdowns, exact
 * anchors) is generated later, one phase at a time, by
 * `PLANNER_PHASE_EXPANSION_SYSTEM_PROMPT` — which reuses almost all of
 * `PLANNER_SYSTEM_PROMPT`'s hard-won rules, just scoped to one phase's worth
 * of output instead of the whole plan's.
 */
export const PLANNER_COARSE_SYSTEM_PROMPT = `
You are the Planner agent for Cellix, an Excel AI assistant — COARSE PLANNING PASS.

Your ONLY job here is to break a large request into its top-level PHASES — the
major sections of work — NOT the detailed subtasks. A second pass will expand
each phase you name into real, formula-level subtasks; do not do that work here.

Output schema:
{
  "phases": [
    {
      "id": "p1",
      "kind": "Short description of what this phase covers, e.g. 'Create one sheet per month with the standard column headers'",
      "targetSheet": "January",
      "dependsOn": [],
      "repeatFor": ["January", "February", ..., "December"]
    }
  ],
  "clarificationsNeeded": [],
  "confidence": "high",
  "reasoning": "One sentence on how you split this."
}

Rules:
- Return ONLY valid JSON — no markdown, no explanation, no commentary.
- A phase is a COHERENT unit of work that will become its OWN, separately-generated batch of subtasks — not a single action. "Create the 12 month sheets" is one phase; "write the KPI row" is a different phase from "write the totals table" only if they are independently useful units (usually they are NOT — prefer fewer, larger phases over many tiny ones).
- ONE SHEET'S BUILD IS ONE PHASE (critical — this is a real, observed failure, not a hypothetical): each phase is expanded by a SEPARATE call that cannot see what any OTHER phase's expansion actually wrote — only that phase's own subtask ids, never its content. Splitting a single cohesive sheet's construction into multiple phases (e.g. "consolidate the data into Main" as phase 2, "build the dashboard/KPIs on Main" as phase 3) means BOTH expansion calls independently decide they need the same underlying structure — a totals table, a KPI band — and BOTH build one from scratch, producing duplicate, colliding writes to the SAME cells. This already happened in production: two phases named "Main sheet" work differently ("consolidates all details" vs "dashboard with KPIs") and both independently rebuilt an identical Monthly Totals table at Main!A4:D16, both tried to create the Main sheet itself, and the second phase's writes were rejected by the overwrite guard as duplicates of the first. NEVER split ONE target sheet's construction (a dashboard, a summary sheet, any single-sheet build with several visual sections — a title, a KPI band, a totals table, a chart, a consolidated-data section) across more than one phase, no matter how many distinct-sounding sections it has. Only split by phase when the pieces target GENUINELY DIFFERENT sheets (month sheets vs. Main vs. Lists) or are a repeated structure (repeatFor).
- Use "repeatFor" ONLY when a phase is the SAME structure applied once per named entry (12 near-identical month sheets is the textbook case). Do not use it for phases that are not a repeated pattern — omit it entirely rather than setting an empty array.
- "targetSheet" for a repeated phase names the FIRST entry as a representative example (e.g. "January" for a 12-month repeatFor) — the expansion pass resolves the real per-entry sheet names from "repeatFor".
- "dependsOn" is PHASE ids, not subtask ids — name another phase here only when THIS phase's work genuinely cannot be planned/described without that phase's sheets/structure already existing (e.g. a dashboard totals phase depends on the month-sheets phase that creates the sheets it will sum).
- Keep phases to the minimum count that keeps each one small and independent — typically 3-8 for even a large multi-sheet build. Do not split one coherent unit into multiple phases just to make more of them.
- If the prompt is small/simple enough that it would never risk truncation as one plan, you may return a SINGLE phase covering everything — this pass exists for large requests, not to force decomposition on small ones.
- confidence = "low" if you are guessing at the user's overall intent (not per-phase detail — that is the expansion pass's concern).
- clarificationsNeeded is a LAST RESORT — same rule as full planning: if you can identify phases at all, do so; only block here if the request is fundamentally unreadable (empty workbook context, no discernible intent).
`;

export function buildCoarsePlannerUserMessage(
  prompt: string,
  context: WorkbookContext,
  history: { role: string; content: string }[],
  promptContext?: string,
): string {
  const workbookSection = promptContext?.trim()
    ? promptContext.trim()
    : [
        `Active sheet: ${context.activeSheetName}`,
        `Sheets: ${context.sheets.map((s) => `${s.name} (${s.rowCount}x${s.columnCount})`).join(', ')}`,
      ].join('\n');

  return `
Conversation history:
${history.map((h) => `${h.role}: ${h.content}`).join('\n')}

User prompt: "${prompt}"

Workbook context:
${workbookSection}

Return JSON only — phases, not subtasks.
`;
}

/**
 * Pass 2 — expand ONE phase into real subtasks. Reuses the base rule text a
 * caller passes in (the existing `PLANNER_SYSTEM_PROMPT` body, minus its
 * output-schema preamble) so every hard-won rule (formula syntax, dropdown
 * validation, number-format preservation, the yearly-ledger anchor math)
 * still applies — this pass is scoped, not a rewrite of what a good plan
 * looks like.
 */
export function buildPhaseExpansionSystemPrompt(baseRules: string): string {
  return `
You are the Planner agent for Cellix, an Excel AI assistant — PHASE EXPANSION PASS.

You are expanding ONE PHASE of a larger plan into real, executable subtasks.
Other phases exist before and after this one — you are told which phase ids
this one depends on and, where already known, the REAL subtask ids those
phases produced. Reference those ids in this phase's own "dependsOn" fields
exactly as given; do not invent new ids for other phases' work.

Output schema (identical to full planning, but ONLY for this phase's subtasks):
{
  "subtasks": [
    { "id": "s1", "description": "...", "targetSheet": "...", "dependsOn": [], "estimatedActions": 3 }
  ],
  "clarificationsNeeded": [],
  "confidence": "high",
  "reasoning": "..."
}

Subtask ids only need to be unique WITHIN this phase's own output (s1, s2, ...) —
they will be namespaced automatically before being merged with other phases'.

${baseRules}
`;
}

export function buildPhaseExpansionUserMessage(opts: {
  originalPrompt: string;
  context: WorkbookContext;
  history: { role: string; content: string }[];
  promptContext?: string;
  phase: { kind: string; targetSheet: string; repeatFor?: string[] };
  /** Real subtask ids already produced for phases this one depends on, so
   *  cross-phase dependsOn can reference them directly. */
  dependencySubtaskIds: string[];
}): string {
  const repeatSection = opts.phase.repeatFor?.length
    ? `\nThis phase repeats the SAME structure once per entry: ${opts.phase.repeatFor.join(', ')}. ` +
      `Produce one group of subtasks per entry — do not collapse them into a single subtask covering all entries, ` +
      `and do not merge more than one entry's work into one subtask (that is exactly the truncation risk this split exists to avoid).`
    : '';

  const dependencySection = opts.dependencySubtaskIds.length
    ? `\nSubtask ids already produced by phases this one depends on (use these directly in "dependsOn" where this phase's own work needs them): ${opts.dependencySubtaskIds.join(', ')}`
    : '';

  const workbookSection = opts.promptContext?.trim()
    ? opts.promptContext.trim()
    : [
        `Active sheet: ${opts.context.activeSheetName}`,
        `Sheets: ${opts.context.sheets.map((s) => `${s.name} (${s.rowCount}x${s.columnCount})`).join(', ')}`,
      ].join('\n');

  return `
Conversation history:
${opts.history.map((h) => `${h.role}: ${h.content}`).join('\n')}

Original user prompt (for context — you are only expanding ONE PHASE of the overall plan it produced): "${opts.originalPrompt}"

This phase: ${opts.phase.kind}
Primary target sheet: ${opts.phase.targetSheet}${repeatSection}${dependencySection}

Workbook context:
${workbookSection}

Return JSON only — this phase's subtasks.
`;
}
