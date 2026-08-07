import { Action } from '../types/agent.types';

/**
 * True when the user clearly asked to clear / empty / blank cell or column values.
 * These are replace intents — overwrite is the goal, not an accident.
 */
export function isClearOrEmptyIntent(message: string): boolean {
  const text = String(message ?? '').toLowerCase();
  if (!text.trim()) return false;

  return (
    /\b(make|set|leave)\b[\s\S]{0,40}\b(empty|blank)\b/i.test(text) ||
    /\b(empty|blank)\b[\s\S]{0,40}\b(column|cells?|values?|remarks?)\b/i.test(text) ||
    /\b(clear|wipe|remove|delete)\b[\s\S]{0,40}\b(values?|contents?|data|remarks?|cells?|column)\b/i.test(
      text,
    ) ||
    /\b(clear|wipe|empty|blank)\b[\s\S]{0,20}\b(the\s+)?remarks?\b/i.test(text) ||
    /\bno\s+values\b/i.test(text)
  );
}

function isBlankWriteValue(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

/**
 * Mark clear/empty writes as explicit overwrite so the frontend guard allows them.
 * Only blanking writes / clear actions are confirmed — not arbitrary replacements.
 */
export function annotateClearIntentOverwrite<T extends Action>(
  actions: T[],
  message: string,
): T[] {
  if (!isClearOrEmptyIntent(message) || actions.length === 0) {
    return actions;
  }

  return actions.map((action) => {
    if (
      action.type === 'CLEAR_CONTENT' ||
      action.type === 'CLEAR_ALL' ||
      action.type === 'CLEAR_CELL' ||
      action.type === 'CLEAR_FORMAT'
    ) {
      return { ...action, explicitOverwriteConfirmed: true };
    }

    if (action.type === 'SET_MATCHING_ROWS' && isBlankWriteValue(action.value)) {
      return { ...action, explicitOverwriteConfirmed: true };
    }

    if (action.type === 'SET_CELL' && isBlankWriteValue(action.value)) {
      return { ...action, explicitOverwriteConfirmed: true };
    }

    if (action.type === 'BATCH_SET' && Array.isArray(action.operations)) {
      const allBlank = action.operations.every(
        (op) => op.formula === undefined && isBlankWriteValue(op.value),
      );
      if (allBlank) {
        return { ...action, explicitOverwriteConfirmed: true };
      }
    }

    return action;
  });
}
