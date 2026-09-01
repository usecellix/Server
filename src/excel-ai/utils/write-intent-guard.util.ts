/**
 * Deterministic write-intent detection — safety net ahead of (and after) route
 * classification so mutation requests never silently fall through to a read-only path.
 */

const WRITE_INTENT_VERBS =
  /\b(sort|filter|delete|remove|insert|add|copy|move|bold|highlight|color|colour|format|merge|split|fill|clear|rename|hide|unhide|freeze|protect|create|build|generate|apply|replace|update|change|set)\b/i;

/**
 * Requests that LOOK like write-intent verbs but are actually questions about
 * the sheet, not mutations of it — must not trip the guard.
 * Note: bare "which" is too broad ("which all month sheets include … columns").
 */
const READ_INTENT_OVERRIDE =
  /\b(what|how many|show me|explain|why|can you tell me|is there)\b/i;

const READ_WHICH_QUESTION =
  /\bwhich\s+(is|are|was|were|of|one|ones|columns?|rows?|sheets?|cells?|values?)\b/i;

/**
 * Compound "ask then mutate" — still write-intent despite a read-style opener.
 *
 * The verb must be IMPERATIVE. F1 (2026-08-27): "...how many invoices are pending
 * payment and what they add up to" matched on "and ... add" and routed a pure
 * question to the write planner, which proposed writing KPI formulas into N2:O5.
 * Requiring "and/then" to be followed directly by the verb (optionally via "also"
 * or "please") keeps "and then delete the column" while rejecting "and what they
 * add up to", where the verb belongs to a subordinate clause.
 */
const READ_THEN_WRITE =
  /\b(?:and|then)\s+(?:also\s+|please\s+|then\s+)?(sort|delete|add|highlight|create|build)\b/i;

/**
 * Verb occurrences that are part of a question, not a command. Checked before the
 * verb list so a noun-phrase or idiomatic use cannot trip write intent.
 * e.g. "what they add up to", "which rows add up", "does it add up".
 */
const VERB_IN_QUESTION_IDIOM = /\badds?\s+up\b|\badding\s+up\b/i;

/**
 * Multi-sheet / yearly ledger / main dashboard scaffold without explicit create/build verbs.
 * e.g. "I like to have multiple sheets for all months… main sheet… dashboard… record payments"
 */
export function isWorkbookScaffoldIntent(message: string): boolean {
  const multiSheet =
    /\b(multiple\s+sheets|monthly\s+sheets|sheets?\s+for\s+(all\s+)?(the\s+)?months?|one\s+sheet\s+per\s+month|sheet\s+for\s+each\s+month|jan(?:uary)?\s*[-–to]+\s*dec(?:ember)?)\b/i.test(
      message,
    ) ||
    /\b(12|twelve)\s+(monthly\s+)?sheets?\b/i.test(message) ||
    /\ball\s+months?\s+in\s+a\s+year\b/i.test(message);

  const mainOrDashboard =
    /\bmain\s+sheet\b/i.test(message) ||
    /\bdashboard\b/i.test(message) ||
    /\bsummary\s+sheet\b/i.test(message);

  const recordDomain =
    /\brecord\s+(payments?|bookings?|guests?)\b/i.test(message) ||
    /\bpayment\s+status\b/i.test(message) ||
    (/\bcheck[\s-]?in\b/i.test(message) && /\bcheck[\s-]?out\b/i.test(message));

  if (multiSheet && (mainOrDashboard || recordDomain)) {
    return true;
  }
  if (mainOrDashboard && recordDomain && /\bsheets?\b/i.test(message)) {
    return true;
  }
  if (
    multiSheet &&
    /\b(include|includes|columns?|headers?)\b/i.test(message) &&
    (/\bi\s+(like|want|need)\s+to\s+have\b/i.test(message) || /\b(need|want)\b/i.test(message))
  ) {
    return true;
  }
  return false;
}

export function hasWriteIntent(message: string): boolean {
  // Scaffold builds always mean mutation, even with soft phrasing or "which … include".
  if (isWorkbookScaffoldIntent(message)) {
    return true;
  }

  // A read-style question wins over an incidental verb match. Checked BEFORE
  // READ_THEN_WRITE so "what ... and what they add up to" stays a question (F1).
  const looksLikeQuestion =
    READ_INTENT_OVERRIDE.test(message) || READ_WHICH_QUESTION.test(message);
  if (looksLikeQuestion && VERB_IN_QUESTION_IDIOM.test(message)) {
    return false;
  }

  if (READ_THEN_WRITE.test(message)) {
    return true;
  }

  if (looksLikeQuestion) {
    return false;
  }

  return WRITE_INTENT_VERBS.test(message);
}

/** Verbs covered by Spec 11 WRITE_INTENT_VERBS (aligned with Spec 01 catalog mutations). */
export const WRITE_INTENT_CATALOG_VERBS = [
  'sort',
  'filter',
  'delete',
  'remove',
  'insert',
  'add',
  'copy',
  'move',
  'bold',
  'highlight',
  'color',
  'colour',
  'format',
  'merge',
  'split',
  'fill',
  'clear',
  'rename',
  'hide',
  'unhide',
  'freeze',
  'protect',
  'create',
  'build',
  'generate',
  'apply',
  'replace',
  'update',
  'change',
  'set',
] as const;
