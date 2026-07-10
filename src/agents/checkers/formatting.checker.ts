import { Injectable } from '@nestjs/common';
import { Action, WorkbookContext } from '../types/agent.types';
import {
  buildCheckerResult,
  CheckerResult,
  SubtaskActionSlice,
} from './checker.types';

const ROW_FORMAT_ACTIONS = new Set(['ADD_ROW', 'INSERT_ROW', 'SET_CELL', 'SET_FORMULA']);

function isBlank(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function isDateLike(value: unknown): boolean {
  if (value instanceof Date) return true;
  if (typeof value !== 'string') return false;
  const text = value.trim();
  return (
    /^\d{4}-\d{2}-\d{2}$/.test(text) ||
    /^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$/.test(text) ||
    /^\d{1,2}[-\s][A-Za-z]{3,9}[-\s]\d{2,4}$/.test(text)
  );
}

function isCurrencyFormat(format: string): boolean {
  const lower = format.toLowerCase();
  return lower.includes('$') || lower.includes('₹') || lower.includes('€') || lower.includes('currency');
}

function isDateFormat(format: string): boolean {
  const lower = format.toLowerCase();
  return (
    (lower.includes('y') || lower.includes('d')) &&
    lower.includes('m') &&
    !lower.includes('general')
  );
}

function resolveReferenceRow(action: Action, targetRow: number): number | null {
  if (action.type === 'ADD_ROW') {
    return targetRow > 0 ? targetRow - 1 : null;
  }
  if (action.type === 'INSERT_ROW') {
    return action.position === 'above' ? targetRow + 1 : targetRow > 0 ? targetRow - 1 : null;
  }
  if (action.type === 'SET_CELL' || action.type === 'SET_FORMULA') {
    return targetRow > 0 ? targetRow - 1 : null;
  }
  return null;
}

function getSheetContext(context: WorkbookContext, sheetName?: string) {
  const name = sheetName ?? context.activeSheetName;
  return context.sheets.find((sheet) => sheet.name === name);
}

@Injectable()
export class FormattingChecker {
  /**
   * Deterministic checks mirroring frontend formatGuard inheritance rules:
   * row-append and blank-cell writes should align with adjacent number formats.
   */
  check(states: SubtaskActionSlice[], context: WorkbookContext): CheckerResult {
    const subtaskResults = states.map((state) => {
      const issues = state.actions.flatMap((action, actionIndex) =>
        this.checkAction(action, actionIndex, state.subtask.targetSheet, context),
      );

      return {
        subtaskId: state.subtask.id,
        passed: issues.every((issue) => issue.severity !== 'error'),
        feedback:
          issues.length === 0
            ? 'Formatting checks passed'
            : issues.map((issue) => issue.description).join('; '),
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
  ) {
    if (!ROW_FORMAT_ACTIONS.has(action.type)) {
      return [];
    }

    const sheetName = action.sheetName ?? defaultSheet;
    const sheet = getSheetContext(context, sheetName);
    if (!sheet) {
      return [];
    }

    const issues = [];
    const targetRow = typeof action.row === 'number' ? action.row : sheet.values.length;
    const targetCol = typeof action.col === 'number' ? action.col : 0;
    const refRow = resolveReferenceRow(action, targetRow);

    if (action.type === 'ADD_ROW' || action.type === 'INSERT_ROW') {
      const data = (action.data ?? action.values ?? []) as unknown[];
      const expectedColumns = Math.max(sheet.columnCount, sheet.values[0]?.length ?? 0, 1);
      if (data.length > 0 && data.length < expectedColumns) {
        issues.push({
          severity: 'warning' as const,
          subtaskId: undefined,
          actionIndex,
          description: `${action.type} data has ${data.length} column(s) but sheet has ${expectedColumns}`,
          suggestion: 'Pad row data to match header width or inherit formatting per column',
        });
      }
    }

    if (refRow == null || refRow < 0) {
      return issues;
    }

    const columnsToCheck =
      action.type === 'ADD_ROW' || action.type === 'INSERT_ROW'
        ? Math.max(
            ((action.data ?? action.values ?? []) as unknown[]).length,
            sheet.columnCount,
            1,
          )
        : 1;

    for (let offset = 0; offset < columnsToCheck; offset += 1) {
      const col = targetCol + offset;
      const refFormat = String(sheet.numberFormats[refRow]?.[col] ?? '');
      if (!refFormat || refFormat === 'General') continue;

      const explicitFormat = action.format?.numberFormat;
      if (explicitFormat && explicitFormat !== refFormat && explicitFormat === 'General') {
        issues.push({
          severity: 'error' as const,
          actionIndex,
          description: `${action.type} explicitly resets column ${col + 1} to General but adjacent row uses "${refFormat}"`,
          suggestion: 'Omit format.numberFormat to inherit from the row above (formatGuard behavior)',
        });
        continue;
      }

      const value =
        action.type === 'ADD_ROW' || action.type === 'INSERT_ROW'
          ? (action.data ?? action.values)?.[offset]
          : action.value;

      if (isBlank(value)) continue;

      if (isDateFormat(refFormat) && !isDateLike(value) && typeof value === 'string') {
        issues.push({
          severity: 'warning' as const,
          actionIndex,
          description: `Column ${col + 1} expects a date (${refFormat}) but value "${value}" may not parse correctly`,
          suggestion: 'Use a date value or let formatGuard coerce on apply',
        });
      }

      if (isCurrencyFormat(refFormat) && typeof value === 'string' && /[A-Za-z]{2,}/.test(value)) {
        issues.push({
          severity: 'warning' as const,
          actionIndex,
          description: `Column ${col + 1} uses currency format "${refFormat}" but value "${value}" looks non-numeric`,
          suggestion: 'Provide a numeric value for currency columns',
        });
      }
    }

    return issues;
  }
}
