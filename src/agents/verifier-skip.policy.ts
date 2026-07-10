// cellix_backend/src/agents/verifier-skip.policy.ts

import { SheetAction } from '../excel-ai/types/sheet-actions.types';

/**
 * Problem solved: Verifier runs on EVERY write operation, adding 1–3 seconds.
 * For simple, low-risk, reversible operations the Accept/Reject preview is
 * already a sufficient safety net.
 *
 * Policy: skip Verifier when ALL of the following are true:
 *  1. No destructive actions (DELETE_SHEET, CLEAR_RANGE on large ranges, DELETE_ROW bulk)
 *  2. Single subtask (complexity is low)
 *  3. Executor returned valid JSON on the first attempt (no parse failure)
 *  4. All actions are reversible via the existing reject/revert mechanism
 */

// Actions that are irreversible or high-risk — always run Verifier
const ALWAYS_VERIFY_ACTION_TYPES = new Set([
  'DELETE_SHEET',
  'CLEAR_RANGE',
  'CLEAR_ALL',
  'CLEAR_CONTENT',
  'DELETE_ROW',
  'DELETE_COLUMN',
  'PROTECT_SHEET',
]);

// Max rows a DELETE_ROW action can touch before requiring Verifier
const BULK_DELETE_ROW_THRESHOLD = 5;

export interface VerifierSkipInput {
  actions: SheetAction[];
  subtaskCount: number;
  executorParsedOnFirstAttempt: boolean;
  hasFormulaActions: boolean;
}

export interface VerifierSkipDecision {
  skip: boolean;
  reason: string;
}

export function shouldSkipVerifier(input: VerifierSkipInput): VerifierSkipDecision {
  const { actions, subtaskCount, executorParsedOnFirstAttempt, hasFormulaActions } = input;

  // Complex multi-subtask operations always need verification
  if (subtaskCount > 2) {
    return { skip: false, reason: 'Multi-subtask operation — verification required' };
  }

  // Formula actions are tricky — Verifier catches formula errors
  if (hasFormulaActions) {
    return { skip: false, reason: 'Formula actions present — verification required' };
  }

  // Executor parse failure means something went wrong — don't skip
  if (!executorParsedOnFirstAttempt) {
    return { skip: false, reason: 'Executor required retry — verification required' };
  }

  // Check for destructive action types
  for (const action of actions) {
    if (ALWAYS_VERIFY_ACTION_TYPES.has(action.type)) {
      // Special case: DELETE_ROW is OK for small counts
      if (action.type === 'DELETE_ROW') {
        const rowCount = getDeleteRowCount(action);
        if (rowCount > BULK_DELETE_ROW_THRESHOLD) {
          return {
            skip: false,
            reason: `Bulk DELETE_ROW (${rowCount} rows) — verification required`,
          };
        }
        // Single row delete: allow skip
        continue;
      }
      return {
        skip: false,
        reason: `Destructive action ${action.type} — verification required`,
      };
    }
  }

  // All checks passed — safe to skip Verifier
  return {
    skip: true,
    reason: `Low-risk single-subtask write (${actions.length} actions) — Verifier skipped`,
  };
}

function getDeleteRowCount(action: SheetAction): number {
  // Handle both single row and range-based delete
  if (action.rowNumbers?.length) {
    return action.rowNumbers.length;
  }
  if ((action as SheetAction & { rows?: unknown[] }).rows && Array.isArray((action as SheetAction & { rows?: unknown[] }).rows)) {
    return (action as SheetAction & { rows: unknown[] }).rows.length;
  }
  if (action.endRow != null && action.row != null) {
    return action.endRow - action.row + 1;
  }
  if ((action as SheetAction & { rowStart?: number; rowEnd?: number }).rowStart != null &&
      (action as SheetAction & { rowStart?: number; rowEnd?: number }).rowEnd != null) {
    const rowStart = (action as SheetAction & { rowStart: number }).rowStart;
    const rowEnd = (action as SheetAction & { rowEnd: number }).rowEnd;
    return rowEnd - rowStart + 1;
  }
  return 1;
}
