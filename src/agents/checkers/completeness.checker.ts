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
      const dropped = state?.droppedActions ?? [];
      const issues = [];

      // A dropped action means the Executor tried to do something we could not apply.
      // Without this the subtask looks complete while silently doing less than asked.
      if (dropped.length > 0) {
        const unknownType = dropped.filter((entry) => entry.reason === 'unknown-type');
        const missingFields = dropped.filter((entry) => entry.reason === 'missing-required-fields');
        const descriptionParts: string[] = [];
        const suggestionParts: string[] = [];
        if (unknownType.length > 0) {
          descriptionParts.push(
            `${unknownType.length} unsupported type(s) (${unknownType.map((e) => e.rawType ?? '(no type)').join(', ')})`,
          );
          suggestionParts.push(
            'use a supported action type from the Available action types list',
          );
        }
        if (missingFields.length > 0) {
          descriptionParts.push(
            `${missingFields.length} action(s) missing a required field (${missingFields.map((e) => e.rawType ?? '(no type)').join(', ')})`,
          );
          // BATCH_SET is the only type this currently applies to — name the field directly.
          suggestionParts.push('include a non-empty "operations" array on every BATCH_SET action');
        }
        issues.push({
          severity: 'error' as const,
          subtaskId: subtask.id,
          description: `Subtask "${subtask.description}" emitted ${dropped.length} unusable action(s), discarded: ${descriptionParts.join('; ')}`,
          suggestion: `Re-emit those steps — ${suggestionParts.join('; ')}`,
        });
      }

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
