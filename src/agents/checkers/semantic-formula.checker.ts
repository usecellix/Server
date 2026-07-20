import { Injectable } from '@nestjs/common';
import {
  buildCheckerResult,
  CheckerResult,
  SubtaskActionSlice,
} from './checker.types';
import { Action, SubTask, WorkbookContext } from '../types/agent.types';

export interface SemanticRule {
  kind: 'tax_percent' | 'column_product_percent';
  rate: number;
  /** Column header names that must appear (via letter or name) in the formula */
  requiredColumns: string[];
  description: string;
}

const TAX_RATE_PATTERN =
  /(\d+(?:\.\d+)?)\s*%\s*(?:of|×|x|\*)/i;
const TAX_OF_QTY_UNIT =
  /tax\s*amount.*(?:qty|quantity).*(?:unit\s*price|price)|(?:qty|quantity).*(?:unit\s*price|price).*tax/i;

/**
 * Derive a checkable semantic rule from an explicit request (e.g. Tier 2 #7/#8).
 * Returns null when the prompt is open-ended and LLM verification should apply.
 */
export function deriveSemanticRuleFromPrompt(prompt: string): SemanticRule | null {
  const text = prompt.trim();
  if (!text) return null;

  const rateMatch = TAX_RATE_PATTERN.exec(text);
  if (rateMatch && /tax/i.test(text)) {
    const rate = Number(rateMatch[1]) / 100;
    const requiredColumns: string[] = [];
    if (/\bqty\b|\bquantity\b/i.test(text)) requiredColumns.push('Qty', 'Quantity');
    if (/unit\s*price/i.test(text)) requiredColumns.push('Unit Price');
    else if (/\bprice\b/i.test(text)) requiredColumns.push('Price', 'Unit Price');

    return {
      kind: 'tax_percent',
      rate,
      requiredColumns: requiredColumns.length ? requiredColumns : ['Qty', 'Unit Price'],
      description: `Tax Amount should be ${rateMatch[1]}% of Qty×Unit Price`,
    };
  }

  if (TAX_OF_QTY_UNIT.test(text)) {
    return {
      kind: 'column_product_percent',
      rate: 0.18,
      requiredColumns: ['Qty', 'Quantity', 'Unit Price'],
      description: 'Tax Amount should be 18% of Qty×Unit Price',
    };
  }

  return null;
}

function formulaUsesRate(formula: string, rate: number): boolean {
  const f = formula.replace(/\s+/g, '').toLowerCase();
  const percent = rate * 100;
  const variants = [
    String(rate),
    rate.toFixed(2),
    String(percent),
    `${percent}%`,
    `(${percent}/100)`,
    `${percent}/100`,
  ];
  return variants.some((v) => f.includes(v.replace(/\s+/g, '').toLowerCase()));
}

function headerLetters(
  context: WorkbookContext,
  sheetName: string | undefined,
  names: string[],
): string[] {
  const sheet =
    context.sheets.find((s) => s.name === sheetName) ??
    context.sheets.find((s) => s.name === context.activeSheetName) ??
    context.sheets[0];
  if (!sheet?.values?.[0]) return [];
  const headers = sheet.values[0].map((h) => String(h ?? '').trim().toLowerCase());
  const letters: string[] = [];
  for (const name of names) {
    const idx = headers.findIndex((h) => h === name.trim().toLowerCase());
    if (idx >= 0) {
      letters.push(columnIndexToLetter(idx));
    }
  }
  return letters;
}

function columnIndexToLetter(col: number): string {
  let index = col + 1;
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

function formulaReferencesColumns(formula: string, letters: string[]): boolean {
  if (letters.length === 0) return true;
  const upper = formula.toUpperCase();
  // Require at least one letter from the expected set (product formulas need 2+)
  const hits = letters.filter((letter) =>
    new RegExp(`${letter}\\d+|\\$${letter}\\$\\d+|${letter}\\d+`, 'i').test(upper),
  );
  return hits.length >= Math.min(2, letters.length) || hits.length === letters.length;
}

export function checkFormulaActionAgainstRule(
  action: Action,
  rule: SemanticRule,
  context: WorkbookContext,
): { passed: boolean; reason: string } {
  const formula = typeof action.formula === 'string' ? action.formula : '';
  if (!formula.startsWith('=')) {
    return { passed: false, reason: 'Expected a SET_FORMULA with a leading =' };
  }

  if (!formulaUsesRate(formula, rule.rate)) {
    return {
      passed: false,
      reason: `Formula does not use the expected rate ${rule.rate} (${rule.description})`,
    };
  }

  const letters = headerLetters(context, action.sheetName, rule.requiredColumns);
  if (!formulaReferencesColumns(formula, letters)) {
    return {
      passed: false,
      reason: `Formula does not reference required columns (${rule.requiredColumns.join(', ')})`,
    };
  }

  return { passed: true, reason: 'Formula matches stated domain rule' };
}

@Injectable()
export class SemanticFormulaChecker {
  check(
    originalPrompt: string,
    subtasks: SubTask[],
    states: SubtaskActionSlice[],
    context: WorkbookContext,
  ): CheckerResult {
    const rule = deriveSemanticRuleFromPrompt(originalPrompt);
    if (!rule) {
      return buildCheckerResult(
        subtasks.map((subtask) => ({
          subtaskId: subtask.id,
          passed: true,
          feedback: 'No explicit semantic rule in prompt — skipped',
          issues: [],
        })),
      );
    }

    const subtaskResults = subtasks.map((subtask) => {
      const state = states.find((s) => s.subtask.id === subtask.id);
      const formulaActions = (state?.actions ?? []).filter(
        (a) => a.type === 'SET_FORMULA' || (typeof a.formula === 'string' && a.formula.startsWith('=')),
      );

      if (formulaActions.length === 0) {
        return {
          subtaskId: subtask.id,
          passed: true,
          feedback: 'No formula actions to semantically check',
          issues: [],
        };
      }

      const failures: string[] = [];
      for (const action of formulaActions) {
        const result = checkFormulaActionAgainstRule(action, rule, context);
        if (!result.passed) failures.push(result.reason);
      }

      return {
        subtaskId: subtask.id,
        passed: failures.length === 0,
        feedback:
          failures.length === 0
            ? rule.description
            : failures.join('; '),
        issues: failures.map((description) => ({
          severity: 'error' as const,
          subtaskId: subtask.id,
          description,
          suggestion: `Revise formula to match: ${rule.description}`,
        })),
      };
    });

    return buildCheckerResult(subtaskResults);
  }
}
