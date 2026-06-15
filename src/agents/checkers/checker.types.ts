import { Action, SubTask, VerifierIssue } from '../types/agent.types';

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
}

export function mergeCheckerResults(results: CheckerResult[]): CheckerResult {
  const subtaskResults = results.flatMap((result) => result.subtaskResults);
  const issues = results.flatMap((result) => result.issues);
  const passed = results.every((result) => result.passed);
  const requiresLlmVerification = results.some((result) => result.requiresLlmVerification);

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
