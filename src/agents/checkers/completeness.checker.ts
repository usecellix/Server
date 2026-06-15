import { Injectable } from '@nestjs/common';
import {
  buildCheckerResult,
  CheckerResult,
  SubtaskActionSlice,
} from './checker.types';
import { SubTask } from '../types/agent.types';

@Injectable()
export class CompletenessChecker {
  check(subtasks: SubTask[], states: SubtaskActionSlice[]): CheckerResult {
    const stateById = new Map(states.map((state) => [state.subtask.id, state]));
    const subtaskResults = subtasks.map((subtask) => {
      const state = stateById.get(subtask.id);
      const actions = state?.actions ?? [];
      const issues = [];

      if (actions.length === 0) {
        issues.push({
          severity: 'error' as const,
          subtaskId: subtask.id,
          description: `Subtask "${subtask.description}" produced no actions`,
          suggestion: 'Re-run executor for this subtask or clarify the request',
        });
      } else if (
        subtask.estimatedActions > 1 &&
        actions.length < Math.max(1, Math.floor(subtask.estimatedActions / 2))
      ) {
        issues.push({
          severity: 'warning' as const,
          subtaskId: subtask.id,
          description: `Subtask "${subtask.description}" produced ${actions.length} action(s) but ~${subtask.estimatedActions} were expected`,
          suggestion: 'Verify the subtask fully addresses the planned step',
        });
      }

      return {
        subtaskId: subtask.id,
        passed: issues.every((issue) => issue.severity !== 'error'),
        feedback:
          issues.length === 0
            ? 'Subtask has actions'
            : issues.map((issue) => issue.description).join('; '),
        issues,
      };
    });

    return buildCheckerResult(subtaskResults);
  }
}
