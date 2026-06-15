import { Injectable, Logger } from '@nestjs/common';
import { Action, WorkbookContext } from '../agents/types/agent.types';
import { colIndexToLetter, letterToColIndex } from '../virtual/shadowWorkbook';
import { ShadowWorkbook } from '../virtual/shadowWorkbook.types';
import { FUNCTION_NAMES, parseFormula } from './formula.parser';
import {
  FormulaValidationIssue,
  FormulaValidationResult,
} from './formula.types';

const EXCEL_ERRORS = [
  '#REF!',
  '#DIV/0!',
  '#NAME?',
  '#VALUE!',
  '#N/A',
  '#NULL!',
  '#NUM!',
  '#SPILL!',
];

const EXCEL_LITERALS = new Set(['TRUE', 'FALSE']);

interface ExtractedFormula {
  formula: string;
  sheetName: string;
  cell?: string;
  actionIndex: number;
}

@Injectable()
export class FormulaValidatorService {
  private readonly logger = new Logger(FormulaValidatorService.name);

  validatePreApply(
    actions: Action[],
    context: WorkbookContext,
    defaultSheet?: string,
  ): FormulaValidationResult {
    const issues: FormulaValidationIssue[] = [];
    const formulas = this.extractFormulas(actions, defaultSheet ?? context.activeSheetName);

    for (const entry of formulas) {
      issues.push(...this.validateSyntax(entry));
      issues.push(...this.validateReferences(entry, context));
      issues.push(...this.validateNamedRanges(entry, context));
    }

    const errors = issues.filter((i) => i.severity === 'error');
    return {
      passed: errors.length === 0,
      issues,
      phase: 'pre_apply',
    };
  }

  checkPostApply(
    shadow: ShadowWorkbook,
    actions: Action[],
    context: WorkbookContext,
    defaultSheet?: string,
  ): FormulaValidationResult {
    const issues: FormulaValidationIssue[] = [];

    for (const key of shadow.changedCells) {
      const bang = key.indexOf('!');
      if (bang === -1) continue;
      const sheetName = key.slice(0, bang);
      const address = key.slice(bang + 1);
      const cell = shadow.sheets.get(sheetName)?.cells.get(address);
      if (!cell) continue;

      const display = String((cell.formula || cell.value) ?? '');
      for (const err of EXCEL_ERRORS) {
        if (display.includes(err)) {
          issues.push({
            severity: 'error',
            code: 'POST_EXEC',
            cell: `${sheetName}!${address}`,
            message: `Cell shows Excel error ${err} after simulation`,
            suggestion: 'Fix formula references or inputs for this cell',
          });
        }
      }
    }

    const formulas = this.extractFormulas(actions, defaultSheet ?? context.activeSheetName);
    const postContext = this.shadowAsContext(shadow, context);
    for (const entry of formulas) {
      issues.push(...this.validateReferences(entry, postContext));
    }

    const errors = issues.filter((i) => i.severity === 'error');
    if (errors.length > 0) {
      this.logger.warn(`Post-apply formula validation: ${errors.length} error(s)`);
    }

    return {
      passed: errors.length === 0,
      issues,
      phase: 'post_apply',
    };
  }

  formatFeedback(issues: FormulaValidationIssue[]): string {
    return issues
      .filter((i) => i.severity === 'error')
      .map(
        (i) =>
          `[${i.code}] ${i.message}${i.suggestion ? ` — Suggestion: ${i.suggestion}` : ''}`,
      )
      .join('\n');
  }

  summarizeForVerifier(results: FormulaValidationResult[]): string {
    const allIssues = results.flatMap((r) => r.issues.filter((i) => i.severity === 'error'));
    if (allIssues.length === 0) {
      return 'Deterministic formula validator: all syntax, references, and post-apply checks passed.';
    }
    return `Deterministic formula validator found ${allIssues.length} issue(s):\n${this.formatFeedback(allIssues)}`;
  }

  private extractFormulas(actions: Action[], defaultSheet: string): ExtractedFormula[] {
    const out: ExtractedFormula[] = [];

    actions.forEach((action, actionIndex) => {
      const sheetName = action.sheetName ?? defaultSheet;

      if (action.type === 'SET_FORMULA' && action.formula) {
        const cell =
          action.address ??
          (action.row !== undefined && action.col !== undefined
            ? `${colIndexToLetter(action.col)}${action.row + 1}`
            : undefined);
        out.push({ formula: action.formula, sheetName, cell, actionIndex });
        return;
      }

      if (action.type === 'SET_CELL' && typeof action.value === 'string' && action.value.startsWith('=')) {
        const cell =
          action.address ??
          (action.row !== undefined && action.col !== undefined
            ? `${colIndexToLetter(action.col)}${action.row + 1}`
            : undefined);
        out.push({ formula: action.value, sheetName, cell, actionIndex });
        return;
      }

      if (action.type === 'ADD_ROW' && Array.isArray(action.data)) {
        action.data.forEach((val, colIdx) => {
          if (typeof val === 'string' && val.startsWith('=')) {
            out.push({
              formula: val,
              sheetName,
              cell: `${colIndexToLetter(colIdx)}?`,
              actionIndex,
            });
          }
        });
        return;
      }

      if (action.type === 'BATCH_SET' && Array.isArray(action.operations)) {
        for (const op of action.operations) {
          if (op.formula) {
            out.push({ formula: op.formula, sheetName, cell: op.address, actionIndex });
          } else if (typeof op.value === 'string' && op.value.startsWith('=')) {
            out.push({ formula: op.value, sheetName, cell: op.address, actionIndex });
          }
        }
      }
    });

    return out;
  }

  private validateSyntax(entry: ExtractedFormula): FormulaValidationIssue[] {
    const issues: FormulaValidationIssue[] = [];
    const formula = entry.formula.trim();

    if (!formula.startsWith('=')) {
      issues.push({
        severity: 'error',
        code: 'SYNTAX',
        actionIndex: entry.actionIndex,
        formula,
        cell: entry.cell,
        message: 'Formula must start with =',
        suggestion: 'Prefix the expression with =',
      });
      return issues;
    }

    const body = formula.slice(1);
    let depth = 0;
    for (const ch of body) {
      if (ch === '(') depth += 1;
      if (ch === ')') depth -= 1;
      if (depth < 0) {
        issues.push({
          severity: 'error',
          code: 'SYNTAX',
          actionIndex: entry.actionIndex,
          formula,
          cell: entry.cell,
          message: 'Unbalanced parentheses in formula',
          suggestion: 'Ensure every ( has a matching )',
        });
        return issues;
      }
    }

    if (depth !== 0) {
      issues.push({
        severity: 'error',
        code: 'SYNTAX',
        actionIndex: entry.actionIndex,
        formula,
        cell: entry.cell,
        message: 'Unbalanced parentheses in formula',
        suggestion: 'Ensure every ( has a matching )',
      });
    }

    if (/\(\s*\)/.test(body)) {
      issues.push({
        severity: 'error',
        code: 'SYNTAX',
        actionIndex: entry.actionIndex,
        formula,
        cell: entry.cell,
        message: 'Empty function call in formula',
        suggestion: 'Provide arguments inside parentheses',
      });
    }

    if (/[^A-Za-z0-9_+\-*/^&=<>:"'.,!% \t[\]\\@#$~;]/.test(body)) {
      issues.push({
        severity: 'error',
        code: 'SYNTAX',
        actionIndex: entry.actionIndex,
        formula,
        cell: entry.cell,
        message: 'Formula contains invalid characters',
        suggestion: 'Use only valid Excel formula syntax',
      });
    }

    return issues;
  }

  private validateReferences(
    entry: ExtractedFormula,
    context: WorkbookContext,
  ): FormulaValidationIssue[] {
    const issues: FormulaValidationIssue[] = [];
    const parsed = parseFormula(entry.formula);
    const defaultSheet = entry.sheetName;

    for (const ref of parsed.cellRefs) {
      const sheetName = ref.sheet ?? defaultSheet;
      const sheet = context.sheets.find((s) => s.name === sheetName);
      if (!sheet) {
        issues.push({
          severity: 'error',
          code: 'REFERENCE',
          actionIndex: entry.actionIndex,
          formula: entry.formula,
          cell: entry.cell,
          message: `Reference ${ref.raw} points to unknown sheet "${sheetName}"`,
          suggestion: 'Use a sheet name from the workbook context',
        });
        continue;
      }

      const colIndex = letterToColIndex(ref.column);
      const rowIndex = ref.row - 1;
      if (rowIndex < 0 || rowIndex >= sheet.rowCount || colIndex < 0 || colIndex >= sheet.columnCount) {
        issues.push({
          severity: 'error',
          code: 'REFERENCE',
          actionIndex: entry.actionIndex,
          formula: entry.formula,
          cell: entry.cell,
          message: `Reference ${ref.raw} is outside sheet "${sheetName}" bounds (${sheet.rowCount} rows × ${sheet.columnCount} cols)`,
          suggestion: 'Adjust the reference to an in-bounds cell',
        });
      }
    }

    for (const range of parsed.rangeRefs) {
      const sheetName = range.sheet ?? defaultSheet;
      const sheet = context.sheets.find((s) => s.name === sheetName);
      if (!sheet) {
        issues.push({
          severity: 'error',
          code: 'REFERENCE',
          actionIndex: entry.actionIndex,
          formula: entry.formula,
          cell: entry.cell,
          message: `Range ${range.raw} points to unknown sheet "${sheetName}"`,
          suggestion: 'Use a sheet name from the workbook context',
        });
        continue;
      }

      const startCol = letterToColIndex(range.startCol);
      const endCol = letterToColIndex(range.endCol);
      const startRow = range.startRow - 1;
      const endRow = range.endRow - 1;

      if (
        startRow < 0 ||
        endRow >= sheet.rowCount ||
        startCol < 0 ||
        endCol >= sheet.columnCount
      ) {
        issues.push({
          severity: 'error',
          code: 'REFERENCE',
          actionIndex: entry.actionIndex,
          formula: entry.formula,
          cell: entry.cell,
          message: `Range ${range.raw} extends outside sheet "${sheetName}" bounds`,
          suggestion: 'Shrink the range to fit within the known used area',
        });
      }
    }

    return issues;
  }

  private validateNamedRanges(
    entry: ExtractedFormula,
    context: WorkbookContext,
  ): FormulaValidationIssue[] {
    const issues: FormulaValidationIssue[] = [];
    const knownNames = new Set(context.namedRanges.map((n) => n.name.toUpperCase()));
    const functionSet = new Set(FUNCTION_NAMES.map((f) => f.toUpperCase()));
    const body = entry.formula.slice(1);

    const tokens = body.match(/\b[A-Za-z_][A-Za-z0-9_.]*\b/g) ?? [];
    for (const token of tokens) {
      const upper = token.toUpperCase();
      if (functionSet.has(upper) || EXCEL_LITERALS.has(upper)) continue;
      if (/^[A-Z]{1,3}\d+$/.test(upper)) continue;
      if (knownNames.has(upper)) continue;

      if (/^[A-Z_][A-Z0-9_]*$/i.test(token) && token.length > 1) {
        issues.push({
          severity: 'warning',
          code: 'NAMED_RANGE',
          actionIndex: entry.actionIndex,
          formula: entry.formula,
          cell: entry.cell,
          message: `Identifier "${token}" is not a known named range or function`,
          suggestion: knownNames.size
            ? `Known named ranges: ${[...knownNames].slice(0, 5).join(', ')}`
            : 'Define the named range or use a cell reference instead',
        });
      }
    }

    return issues;
  }

  private shadowAsContext(
    shadow: ShadowWorkbook,
    base: WorkbookContext,
  ): WorkbookContext {
    return {
      ...base,
      activeSheetName: shadow.activeSheetName,
      sheets: base.sheets.map((sheet) => {
        const shadowSheet = shadow.sheets.get(sheet.name);
        if (!shadowSheet) return sheet;
        return {
          ...sheet,
          rowCount: Math.max(sheet.rowCount, shadowSheet.rowCount),
          columnCount: Math.max(sheet.columnCount, shadowSheet.columnCount),
        };
      }),
    };
  }
}
