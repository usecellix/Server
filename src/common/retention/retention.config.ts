/**
 * P0.3 — data-retention limits (§1.5 "no client data retention").
 *
 * Decision taken 2026-11-07 (founder asked to implement every step; this is the
 * option recommended in the plan): audit trail + undo data stay as SERVER METADATA
 * with a hard expiry, never in a workbook hidden sheet, and row-level undo data
 * never leaves the add-in (client/src/engine/rowSnapshot.ts).
 *
 * What still holds cell values server-side, and for how long:
 *  - change sets (before/after values used by Revert) → CHANGESET_TTL_HOURS (default 7 days)
 *  - ai_prompts (the typed prompt)                  → AI_PROMPT_TTL_HOURS (default 24 h)
 *  - conversations (chat history)                   → CONVERSATION_TTL_HOURS (unchanged, 90 days)
 *  - planner.log on disk (full LLM input/output)    → OFF unless PLANNER_FILE_LOG=on
 *  - requests.log (typed message + full response)   → metadata only unless REQUEST_LOG_BODIES=on
 * Removing values entirely would break server-side Revert for every action and is a
 * separate redesign — see CELLIX_TEST_CASES.md (P0.3) for what the sentinel test proves.
 */
const hours = (name: string, fallback: number): number => {
  const raw = Number(process.env[name]);
  return Number.isFinite(raw) && raw > 0 ? raw : fallback;
};

export const RETENTION = {
  changeSetHours: () => hours('CHANGESET_TTL_HOURS', 24 * 7),
  aiPromptHours: () => hours('AI_PROMPT_TTL_HOURS', 24),
  /** requests.log keeps the typed message and the full response (cell values) only when this is on. */
  requestLogBodiesEnabled: () => String(process.env.REQUEST_LOG_BODIES ?? '').trim().toLowerCase() === 'on',
  plannerFileLogEnabled: () => String(process.env.PLANNER_FILE_LOG ?? '').trim().toLowerCase() === 'on',
};

export const expiryFromNow = (hoursToKeep: number): Date => new Date(Date.now() + hoursToKeep * 3_600_000);
