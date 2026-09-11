import { Injectable } from '@nestjs/common';
import { Action, VerifierIssue, WorkbookContext } from '../types/agent.types';
import { sanitizeExcelSheetName } from '../../excel-ai/utils/sheet-name.util';
import {
  buildCheckerResult,
  CheckerResult,
  SubtaskActionSlice,
} from './checker.types';

const SHEET_CREATE_TYPES = new Set(['CREATE_SHEET', 'ADD_SHEET']);

/**
 * Catches a class of failure the other four checkers cannot see by
 * construction: they compare the Executor's output against the Planner's own
 * estimate (`estimatedActions`) or against static rules, never against what
 * the subtask actually asked for. `CompletenessChecker` counts actions;
 * nothing compares their *content* to `subtask.targetSheet`.
 *
 * Concrete failure this closes — recorded verbatim in
 * `COMPETITIVE_STUDY_SHORTCUT.md:71`: a subtask targeting sheet "Main"
 * emitted a CREATE_SHEET that landed as "Main 2" (an idempotent "if it
 * doesn't exist" create that didn't respect the existing sheet). The action
 * is well-formed, the count matches, and every existing checker passes it —
 * this is the first one that reads the emitted sheet name at all.
 *
 * Deliberately conservative: only fires when `subtask.targetSheet` reads as
 * a genuine single sheet name (the Planner sometimes leaves it as a vague
 * area description for non-sheet-creating subtasks), and skips entirely when
 * the sheet already existed before this run — reusing an existing sheet is
 * correct behavior, not a mismatch, and is exactly what an idempotent
 * "if it doesn't exist" create is supposed to do.
 */
@Injectable()
export class StructuralIntentChecker {
  check(states: SubtaskActionSlice[], context: WorkbookContext): CheckerResult {
    const preExistingSheetNames = new Set(
      context.sheets.map((sheet) => sheet.name.trim().toLowerCase()),
    );

    const subtaskResults = states.map((state) => {
      const desc = state.subtask.description ?? '';
      const intendsSheetCreate = /\b(create|add|make)\b.*\bsheet\b/i.test(desc);

      if (!intendsSheetCreate) {
        return {
          subtaskId: state.subtask.id,
          passed: true,
          feedback: 'Structural intent checks skipped (not a sheet-creation flow)',
          issues: [],
        };
      }

      const expectedName = sanitizeExcelSheetName(state.subtask.targetSheet ?? '');
      const issues = this.checkActions(state.actions, expectedName, preExistingSheetNames);

      const passed = issues.every((issue) => issue.severity !== 'error');
      const feedback =
        issues.length === 0
          ? 'Structural intent checks passed'
          : issues.map((issue) => issue.description).join('; ');

      return {
        subtaskId: state.subtask.id,
        passed,
        feedback,
        issues,
      };
    });

    return buildCheckerResult(subtaskResults);
  }

  private checkActions(
    actions: Action[],
    expectedName: string,
    preExistingSheetNames: Set<string>,
  ): VerifierIssue[] {
    const expectedKey = expectedName.trim().toLowerCase();
    if (!expectedKey || preExistingSheetNames.has(expectedKey)) return [];

    const issues: VerifierIssue[] = [];

    actions.forEach((action, actionIndex) => {
      if (!SHEET_CREATE_TYPES.has(action.type)) return;

      const raw = (action as { sheetName?: string; name?: string }).sheetName ?? (action as { name?: string }).name;
      if (!raw) return;

      const actualName = sanitizeExcelSheetName(raw);
      if (actualName.trim().toLowerCase() === expectedKey) return;

      issues.push({
        severity: 'error',
        subtaskId: undefined,
        actionIndex,
        description:
          `Subtask targets sheet "${expectedName}" but the emitted ${action.type} action ` +
          `creates a sheet named "${actualName}" instead.`,
        suggestion: `Emit ${action.type} with sheetName/name exactly "${expectedName}" — ` +
          `do not append a suffix or alter the requested name.`,
      });
    });

    return issues;
  }
}
