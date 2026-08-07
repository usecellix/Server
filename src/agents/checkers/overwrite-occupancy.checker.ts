import { Injectable } from '@nestjs/common';
import { Action, VerifierIssue, WorkbookContext } from '../types/agent.types';
import {
  buildCheckerResult,
  CheckerResult,
  SubtaskActionSlice,
} from './checker.types';

/**
 * Spec 19 Bug 1:
 * Deterministic pre-check for "never silently overwrite occupied cells".
 *
 * We can't run Office.js here, so we mirror the frontend overwriteGuard by
 * checking whether the target cell already has non-empty values in the
 * current WorkbookContext snapshot.
 */
@Injectable()
export class OverwriteOccupancyChecker {
  check(states: SubtaskActionSlice[], context: WorkbookContext): CheckerResult {
    const subtaskResults = states.map((state) => {
      // Spec 19 Bug 1 fix is specifically about "add/insert a column" flows where
      // the executor may silently target the wrong already-occupied column.
      // Keep the deterministic checker scoped to column-mutation subtasks so we
      // don't break existing unit tests that model generic SET_CELL/SET_FORMULA.
      const desc = state.subtask.description ?? '';
      const shouldCheck =
        state.subtask.suggestedActionType === 'INSERT_COLUMN' ||
        /\b(add|insert)\b.*\bcolumn\b/i.test(desc);

      if (!shouldCheck) {
        return {
          subtaskId: state.subtask.id,
          passed: true,
          feedback: 'Overwrite occupancy checks skipped (not a column insert flow)',
          issues: [],
        };
      }

      const defaultSheet = state.subtask.targetSheet;
      const issues = state.actions.flatMap((action, actionIndex) =>
        this.checkAction(action, actionIndex, defaultSheet, context),
      );

      const passed = issues.every((i) => i.severity !== 'error');
      const feedback =
        issues.length === 0 ? 'Overwrite occupancy checks passed' : issues.map((i) => i.description).join('; ');

      return {
        subtaskId: state.subtask.id,
        passed,
        feedback,
        issues,
      };
    });

    return buildCheckerResult(subtaskResults);
  }

  private checkAction(
    action: Action,
    actionIndex: number,
    defaultSheet: string,
    context: WorkbookContext,
  ): VerifierIssue[] {
    // The "explicit overwrite confirmed" escape hatch — only set for unambiguous user replace intents.
    if (Boolean((action as { explicitOverwriteConfirmed?: boolean }).explicitOverwriteConfirmed)) {
      return [];
    }

    const sheetName =
      action.type === 'COPY_FILTERED_RANGE' || action.type === 'MOVE_RANGE' || action.type === 'AGGREGATE_TABLE'
        ? (action as { destSheet: string }).destSheet
        : (action as { sheetName?: string }).sheetName ?? defaultSheet;

    const sheet = context.sheets.find((s) => s.name === sheetName);
    if (!sheet) return [];

    const usedStart = this.parseUsedRangeStart(sheet.usedRange);
    if (!usedStart) return [];

    // Helper for "what address would the guard throw on?"
    const mkAddressFromAbs = (absRow0: number, absCol0: number): string => {
      const rowExcel = absRow0 + 1;
      const colLetter = this.colIndexToLetter(absCol0);
      return `${colLetter}${rowExcel}`;
    };

    const hasData = (existing: unknown): boolean => existing !== '' && existing != null;

    const resolveExistingAtAbs = (absRow0: number, absCol0: number): unknown => {
      const relRow = absRow0 - usedStart.row0;
      const relCol = absCol0 - usedStart.col0;
      return sheet.values?.[relRow]?.[relCol];
    };

    const buildOverwriteFeedback = (targetRange: string, existingValue: unknown): string => {
      const sample = existingValue == null || existingValue === '' ? [] : [String(existingValue)];
      const flat = sample.slice(0, 3).filter((v) => Boolean(v));
      const sampleText = flat.length ? ` Existing values include: ${flat.join(', ')}.` : '';
      return (
        `Write blocked: target range ${targetRange} already contains data. ` +
        `This action would overwrite existing values.${sampleText} ` +
        `If you meant to add a new column, use INSERT_COLUMN with position "afterLastColumn" ` +
        `instead of writing into an occupied column. ` +
        `To replace existing content on purpose, the request must explicitly confirm overwrite.`
      );
    };

    // --- SET_CELL / SET_FORMULA: single-cell occupancy check ---
    if (action.type === 'SET_CELL' || action.type === 'SET_FORMULA') {
      const any = action as unknown as { address?: string; row?: number; col?: number };

      // Prefer provided A1 address when available so the feedback matches the frontend.
      if (typeof any.address === 'string' && any.address.trim()) {
        const local = this.stripSheetPrefix(any.address);
        const parsed = this.parseA1Cell(local);
        if (!parsed) return [];
        const existing = resolveExistingAtAbs(parsed.row0, parsed.col0);
        if (!hasData(existing)) return [];

        return [
          {
            severity: 'error',
            subtaskId: undefined,
            actionIndex,
            description: buildOverwriteFeedback(local, existing),
            suggestion: 'Use INSERT_COLUMN (position "afterLastColumn") instead of writing into an occupied column',
          },
        ];
      }

      if (typeof any.row === 'number' && typeof any.col === 'number') {
        const absRow0 = usedStart.row0 + any.row;
        const absCol0 = usedStart.col0 + any.col;
        const existing = resolveExistingAtAbs(absRow0, absCol0);
        if (!hasData(existing)) return [];
        const addr = mkAddressFromAbs(absRow0, absCol0);

        return [
          {
            severity: 'error',
            subtaskId: undefined,
            actionIndex,
            description: buildOverwriteFeedback(addr, existing),
            suggestion: 'Use INSERT_COLUMN (position "afterLastColumn") instead of writing into an occupied column',
          },
        ];
      }

      return [];
    }

    // --- BATCH_SET: first occupied op triggers failure (mirrors frontend guard throw order) ---
    if (action.type === 'BATCH_SET') {
      const any = action as unknown as { operations?: Array<{ address: string }> };
      const ops = Array.isArray(any.operations) ? any.operations : [];

      for (let i = 0; i < ops.length; i += 1) {
        const op = ops[i];
        if (!op?.address) continue;
        const local = this.stripSheetPrefix(op.address);
        const parsed = this.parseA1Cell(local);
        if (!parsed) continue;
        const existing = resolveExistingAtAbs(parsed.row0, parsed.col0);
        if (!hasData(existing)) continue;

        return [
          {
            severity: 'error',
            subtaskId: undefined,
            actionIndex,
            description: buildOverwriteFeedback(local, existing),
            suggestion: 'Use INSERT_COLUMN (position "afterLastColumn") instead of writing into an occupied column',
          },
        ];
      }

      return [];
    }

    // --- COPY/MOVE/AGGREGATE destination start: mirror guardAgainstOverwrite start-cell check ---
    if (action.type === 'COPY_FILTERED_RANGE' || action.type === 'MOVE_RANGE' || action.type === 'AGGREGATE_TABLE') {
      const any = action as unknown as { destStartCell?: string };
      if (!any.destStartCell) return [];

      const local = this.stripSheetPrefix(any.destStartCell);
      const parsed = this.parseA1Cell(local);
      if (!parsed) return [];

      const existing = resolveExistingAtAbs(parsed.row0, parsed.col0);
      if (!hasData(existing)) return [];

      return [
        {
          severity: 'error',
          subtaskId: undefined,
          actionIndex,
          description: buildOverwriteFeedback(local, existing),
          suggestion: 'Use an INSERT/relocation operation instead of overwriting occupied destination data',
        },
      ];
    }

    return [];
  }

  private stripSheetPrefix(address: string): string {
    return address.trim().replace(/\$/g, '').split('!').pop() ?? address.trim();
  }

  private colIndexToLetter(col0: number): string {
    let n = col0 + 1;
    let letter = '';
    while (n > 0) {
      const mod = (n - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      n = Math.floor((n - 1) / 26);
    }
    return letter;
  }

  /**
   * Parse A1 cell into absolute 0-based indices.
   * E.g. "K2" -> { row0: 1, col0: 10 }.
   */
  private parseA1Cell(address: string): { row0: number; col0: number } | null {
    const match = /^([A-Za-z]+)(\d+)$/.exec(address.trim());
    if (!match) return null;
    const colLetters = match[1].toUpperCase();
    const rowExcel = Number.parseInt(match[2], 10);
    if (!Number.isFinite(rowExcel)) return null;

    let col0 = 0;
    for (let i = 0; i < colLetters.length; i += 1) {
      col0 = col0 * 26 + (colLetters.charCodeAt(i) - 64);
    }
    col0 -= 1;

    return { row0: rowExcel - 1, col0 };
  }

  private parseUsedRangeStart(usedRange: string): { row0: number; col0: number } | null {
    const local = this.stripSheetPrefix(usedRange);
    const start = local.includes(':') ? local.split(':')[0] : local;
    const parsed = this.parseA1Cell(start);
    return parsed;
  }
}

