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
  check(
    states: SubtaskActionSlice[],
    context: WorkbookContext,
    /**
     * Sheets an EARLIER wave of this same run already created — TASKS.md #294.
     * Those are legitimately present by the time the step that fills them
     * runs, and are invisible here otherwise: the checker only ever grades
     * this wave's own states.
     */
    sheetsCreatedByEarlierWaves: Set<string> = new Set(),
  ): CheckerResult {
    const preExistingSheetNames = new Set(
      context.sheets.map((sheet) => sheet.name.trim().toLowerCase()),
    );

    /**
     * Sheets some OTHER subtask genuinely created with a real create action —
     * TASKS.md #294. A sheet built by an earlier wave (the deterministic
     * header step, say) is legitimately already there by the time the step
     * that fills it runs, and demanding a second create would be wrong.
     *
     * Computed from the states' own actions rather than the passed context on
     * purpose: `virtualApply` conjures a shadow sheet the moment ANYTHING
     * writes to it, so a context enriched from the shadow cannot tell a sheet
     * that was created from one that was merely written to — which is exactly
     * how a subtask that wrote Main's title without ever creating Main passed
     * this check in a live run.
     */
    const createdByOthers = new Map<string, Set<string>>();
    for (const state of states) {
      for (const action of state.actions) {
        if (!SHEET_CREATE_TYPES.has(action.type)) continue;
        const raw =
          (action as { sheetName?: string; name?: string }).name ??
          (action as { sheetName?: string }).sheetName;
        if (!raw) continue;
        const key = sanitizeExcelSheetName(raw).trim().toLowerCase();
        const owners = createdByOthers.get(key) ?? new Set<string>();
        owners.add(state.subtask.id);
        createdByOthers.set(key, owners);
      }
    }

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
      // A create by ANY other subtask counts; one by this subtask itself does
      // not, or the check would be vacuous.
      const owners = createdByOthers.get(expectedName.trim().toLowerCase());
      const createdElsewhere = Boolean(
        owners && [...owners].some((id) => id !== state.subtask.id),
      );
      const key = expectedName.trim().toLowerCase();
      const alreadyThere =
        createdElsewhere || sheetsCreatedByEarlierWaves.has(key)
          ? new Set([...preExistingSheetNames, key])
          : preExistingSheetNames;
      const issues = this.checkActions(state.actions, expectedName, alreadyThere);

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
    let foundMatchingCreate = false;

    actions.forEach((action, actionIndex) => {
      if (!SHEET_CREATE_TYPES.has(action.type)) return;

      const raw = (action as { sheetName?: string; name?: string }).sheetName ?? (action as { name?: string }).name;
      if (!raw) return;

      const actualName = sanitizeExcelSheetName(raw);
      if (actualName.trim().toLowerCase() === expectedKey) {
        foundMatchingCreate = true;
        return;
      }

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

    // TASKS.md #266 — the loop above only ever catches a WRONG-NAMED create;
    // a subtask whose description says "create sheet X" but whose actions
    // contain no create at all (writes to X went out, the create itself just
    // never got emitted) passed silently. Live: a 12-subtask batched
    // Executor call emitted every month's writes but dropped one month's
    // ADD_SHEET/CREATE_SHEET outright — same class of failure as #262, just
    // at generation time instead of planning time, so #262's plan-level net
    // could not see it (the subtask's OWN description already promised the
    // create; nothing was missing from the plan). Skip when a wrong-named
    // issue already fired above — that is a different, already-reported
    // problem, and flagging "no create" on top of it would just be noise.
    if (!foundMatchingCreate && issues.length === 0) {
      issues.push({
        severity: 'error',
        subtaskId: undefined,
        actionIndex: undefined,
        description:
          `Subtask targets sheet "${expectedName}" and its own description says to create it, ` +
          `but the emitted actions contain no ADD_SHEET/CREATE_SHEET for it at all.`,
        suggestion: `Emit an ADD_SHEET (or CREATE_SHEET) action with sheetName/name exactly "${expectedName}" ` +
          `before any of this subtask's writes to that sheet.`,
      });
    }

    return issues;
  }
}
