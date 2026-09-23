import { Injectable } from '@nestjs/common';
import { Action, VerifierIssue } from '../types/agent.types';
import { buildCheckerResult, CheckerResult, SubtaskActionSlice } from './checker.types';
import { findHeaderMismatches } from '../utils/build-spec.util';

/**
 * A sheet whose header row does not match what the user asked for —
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 1.
 *
 * Every other checker grades what was emitted against the plan; none compares
 * it with the user's own words. A live run shipped sheets with Excel's
 * placeholder headers ("Column1"...) and others with headers unlike the
 * prompt's list, and each passed.
 *
 * Only subtasks carrying `expectedHeaders` are judged, so a prompt with no
 * explicit column list is never affected. Extra columns are allowed (computed
 * ones like "Nights"); a missing, renamed or reordered one fails.
 */

const columnIndex = (letters: string): number =>
  letters
    .toUpperCase()
    .split('')
    .reduce((total, char) => total * 26 + (char.charCodeAt(0) - 64), 0);

const ROW_ONE_CELL = /^\$?([A-Za-z]{1,3})\$?1$/;

interface HeaderCell {
  col: number;
  text: string;
}

/** Row-1 string cells this subtask wrote to `sheet`, whichever action shape carried them. */
function collectRowOneCells(actions: Action[], sheet: string): HeaderCell[] {
  const target = sheet.trim().toLowerCase();
  const cells: HeaderCell[] = [];

  for (const action of actions) {
    const record = action as unknown as Record<string, unknown>;
    if (String(record.sheetName ?? '').trim().toLowerCase() !== target) continue;

    if (action.type === 'BATCH_SET' && Array.isArray(record.operations)) {
      for (const op of record.operations as Array<Record<string, unknown>>) {
        const match = ROW_ONE_CELL.exec(String(op.address ?? ''));
        if (match && typeof op.value === 'string' && !op.formula) {
          cells.push({ col: columnIndex(match[1]), text: op.value });
        }
      }
      continue;
    }

    const anchor = String(record.startCell ?? record.address ?? record.range ?? '');
    const anchorMatch = /^\$?([A-Za-z]{1,3})\$?1(?::|$)/.exec(anchor);
    if (anchorMatch && Array.isArray(record.values)) {
      const first = record.values[0];
      const row = Array.isArray(first) ? first : record.values;
      const startCol = columnIndex(anchorMatch[1]);
      row.forEach((value, offset) => {
        if (typeof value === 'string') cells.push({ col: startCol + offset, text: value });
      });
    }
  }

  return cells.sort((a, b) => a.col - b.col);
}

/** True when the subtask created a header-bearing table on `sheet`. */
function createsHeaderedTable(actions: Action[], sheet: string): boolean {
  const target = sheet.trim().toLowerCase();
  return actions.some((action) => {
    const record = action as unknown as Record<string, unknown>;
    return (
      action.type === 'CREATE_TABLE' &&
      String(record.sheetName ?? '').trim().toLowerCase() === target &&
      record.hasHeaders !== false
    );
  });
}

@Injectable()
export class SpecConformanceChecker {
  check(states: SubtaskActionSlice[]): CheckerResult {
    const subtaskResults = states.map((state) => {
      const expected = state.subtask.expectedHeaders;
      const issues: VerifierIssue[] = [];

      if (expected && expected.length > 0) {
        const sheet = state.subtask.targetSheet;
        const written = collectRowOneCells(state.actions, sheet);

        if (written.length === 0) {
          // Nothing wrote the header row, yet a table over it would be built
          // with Excel's own "Column1".."ColumnN" placeholders.
          if (createsHeaderedTable(state.actions, sheet)) {
            issues.push({
              severity: 'error',
              subtaskId: state.subtask.id,
              description:
                `Sheet "${sheet}" got a table but its header row was never written, so Excel would ` +
                `name the columns "Column1", "Column2"...`,
              suggestion:
                `Write row 1 first, then create the table over it. Headers: ${expected.join(' | ')}`,
            });
          }
        } else {
          const missing = findHeaderMismatches(
            expected,
            written.map((cell) => cell.text),
          );
          if (missing.length > 0) {
            issues.push({
              severity: 'error',
              subtaskId: state.subtask.id,
              description:
                `Sheet "${sheet}" header row does not match the columns the user asked for — ` +
                `missing, renamed or out of order: ${missing.join(', ')}.`,
              suggestion:
                `Row 1 must contain these columns verbatim and in this order (computed columns may be ` +
                `added between or after them): ${expected.join(' | ')}`,
            });
          }
        }
      }

      const passed = issues.length === 0;
      return {
        subtaskId: state.subtask.id,
        passed,
        feedback: passed
          ? 'Spec conformance checks passed'
          : issues.map((issue) => issue.description).join('; '),
        issues,
      };
    });

    return buildCheckerResult(subtaskResults);
  }
}
