import { Action, DroppedAction, SubTask, VerifierIssue } from '../types/agent.types';

export interface CheckerSubtaskResult {
  subtaskId: string;
  passed: boolean;
  feedback: string;
  issues: VerifierIssue[];
}

export interface CheckerResult {
  passed: boolean;
  requiresLlmVerification: boolean;
  feedback: string;
  issues: VerifierIssue[];
  subtaskResults: CheckerSubtaskResult[];
}

export interface SubtaskActionSlice {
  subtask: SubTask;
  actions: Action[];
  /** Actions the Executor emitted that could not be normalized into a usable type. */
  droppedActions?: DroppedAction[];
}

export function mergeCheckerResults(results: CheckerResult[]): CheckerResult {
  // Important: the deterministic pipeline runs multiple checkers per subtask.
  // We must merge them down to ONE outcome per subtask, otherwise retry-locking
  // may consider a subtask "passed" when *some* checker passed even if a later
  // checker failed (and then retry feedback can pick the wrong failure).
  const issues = results.flatMap((result) => result.issues);
  const passed = results.every((result) => result.passed);
  const requiresLlmVerification = results.some((result) => result.requiresLlmVerification);

  const bySubtaskId = new Map<string, CheckerSubtaskResult[]>();
  for (const checkerResult of results) {
    for (const sr of checkerResult.subtaskResults) {
      const arr = bySubtaskId.get(sr.subtaskId) ?? [];
      arr.push(sr);
      bySubtaskId.set(sr.subtaskId, arr);
    }
  }

  const subtaskResults: CheckerSubtaskResult[] = Array.from(bySubtaskId.entries()).map(
    ([subtaskId, subResults]) => {
      const mergedPassed = subResults.every((r) => r.passed);
      const mergedIssues = subResults.flatMap((r) => r.issues);

      // Pick the most actionable failing feedback for retry.
      const failing = subResults.filter((r) => !r.passed);
      const preferredFail =
        failing.find((r) => /Write blocked|INSERT_COLUMN/i.test(r.feedback)) ?? failing[0];
      const feedback = preferredFail?.feedback ?? subResults[0]?.feedback ?? '';

      return {
        subtaskId,
        passed: mergedPassed,
        feedback,
        issues: mergedIssues,
      };
    },
  );

  const feedback = results
    .map((result) => result.feedback)
    .filter(Boolean)
    .join(' ');

  return {
    passed,
    requiresLlmVerification,
    feedback: feedback || (passed ? 'Deterministic checks passed' : 'Deterministic checks failed'),
    issues,
    subtaskResults,
  };
}

export function buildCheckerResult(
  subtaskResults: CheckerSubtaskResult[],
  options?: { forceLlmVerification?: boolean },
): CheckerResult {
  const issues = subtaskResults.flatMap((result) =>
    result.issues.map((issue) => ({
      ...issue,
      subtaskId: issue.subtaskId ?? result.subtaskId,
    })),
  );
  const passed = subtaskResults.every((result) => result.passed);
  const hasWarnings = issues.some((issue) => issue.severity === 'warning');
  const hasErrors = issues.some((issue) => issue.severity === 'error');

  return {
    passed: passed && !hasErrors,
    requiresLlmVerification:
      Boolean(options?.forceLlmVerification) || hasWarnings,
    feedback: passed
      ? hasWarnings
        ? 'Deterministic checks passed with warnings'
        : 'Deterministic checks passed'
      : 'Deterministic checks failed',
    issues,
    subtaskResults,
  };
}
