/**
 * Deterministic write-intent detection — safety net ahead of (and after) route
 * classification so mutation requests never silently fall through to a read-only path.
 */

const WRITE_INTENT_VERBS =
  /\b(sort|filter|delete|remove|insert|add|bold|highlight|color|colour|format|merge|split|fill|clear|rename|hide|unhide|freeze|protect|create|build|generate|apply|replace|update|change|set)\b/i;

/**
 * Requests that LOOK like write-intent verbs but are actually questions about
 * the sheet, not mutations of it — must not trip the guard.
 */
const READ_INTENT_OVERRIDE =
  /\b(what|which|how many|show me|explain|why|can you tell me|is there)\b/i;

/** Compound "ask then mutate" — still write-intent despite a read-style opener. */
const READ_THEN_WRITE =
  /\b(and|then)\b.*\b(sort|delete|add|highlight)\b/i;

export function hasWriteIntent(message: string): boolean {
  if (READ_INTENT_OVERRIDE.test(message) && !READ_THEN_WRITE.test(message)) {
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
