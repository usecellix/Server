import { Injectable } from '@nestjs/common';
import { letterToColIndex } from '../../virtual/shadowWorkbook';
import { buildCheckerResult, CheckerResult, SubtaskActionSlice } from './checker.types';

/**
 * A subtask can silently deviate from a formula spelled out in its own
 * description — TASKS.md #328. Live `run_1790326084342_hctud2t`: p3_s4's
 * description said `B2 =SUM(B5:B16)`; the Executor wrote `=SUM(B5:B10)`
 * instead, the subtask completed, and every existing checker passed it —
 * none of them compares a WRITTEN formula against what the subtask's own
 * text dictated, only whether a formula exists at all (ComputedColumnChecker,
 * #270) or matches a rule derived from the user's original prompt
 * (SemanticFormulaChecker, a different, narrower source of truth). #327
 * removed this exposure for Main specifically by replacing its
 * model-generated subtasks with deterministic code, but month-sheet and
 * other generic subtasks still go through the Executor LLM with spelled-out
 * formulas in their description, and nothing checked them until now.
 *
 * The planner is instructed (planner.prompt.ts's ledger rule (d)) to write
 * the FULL formula in the description "so the Executor transcribes them
 * instead of inventing one" — this checker holds it to exactly that: a cell
 * reference followed by `=<formula>` in the description is a promise, and a
 * SET_FORMULA action at that same cell writing something else is a broken one.
 *
 * Deliberately narrow, matching #270/#294's own precedent for a deterministic
 * checker of this kind:
 * - Only cells the description ITSELF names are checked — a subtask that
 *   spells out three cells and adds a fourth of its own accord is not
 *   flagged for the fourth; that is legitimate elaboration, not deviation.
 * - Comparison is whitespace/case-normalized, not byte-exact — `=SUM(B5:B16)`
 *   and `= sum(b5:b16)` are the same instruction. A genuinely different
 *   formula (different range, different function, different operator) fails.
 * - A cell the description names but the subtask never wrote at all is NOT
 *   this checker's concern — CompletenessChecker/ComputedColumnChecker
 *   already cover "promised but never written"; this is specifically
 *   "written, but wrong".
 */

/** `B2 =` marks the start of a spelled-out formula; the rest is scanned char by char below. */
const CELL_FORMULA_START = /\b([A-Za-z]{1,3})(\d{1,7})\s*=/g;

export interface DescribedFormula {
  row: number;
  col: number;
  cellRef: string;
  formula: string;
}

/**
 * Live planner descriptions are prose with formulas embedded, not a clean
 * list — e.g. "...rows 1-2, with each label directly above its value in the
 * SAME column: A1=… The KPI cells sum the Monthly Totals table's own
 * columns — do not re-derive cross-sheet formulas here. Apply number..."
 * (the exact text from TASKS.md #320's live run). A formula's own content
 * (sheet names, string literals, `.`/`,`/`:`/`-`) overlaps too much with
 * ordinary punctuation to split on a fixed character — this instead reads
 * forward from `=` while parentheses stay balanced, and once they do,
 * stops at the first sign of prose resuming: a `.`/`—`/em-dash, a comma NOT
 * inside a string literal, or two consecutive words that read as English
 * (a lowercase word following whitespace, since a formula's own tokens are
 * refs/functions/numbers/operators, never bare lowercase words outside a
 * quoted string).
 */
function readFormula(text: string, startAt: number): string {
  let i = startAt;
  let depth = 0;
  let inString = false;
  for (; i < text.length; i += 1) {
    const ch = text[i];
    if (inString) {
      if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '(') {
      depth += 1;
      continue;
    }
    if (ch === ')') {
      depth = Math.max(0, depth - 1);
      continue;
    }
    if (depth > 0) continue; // never stop mid-parenthesis
    if (ch === ',' || ch === '.' || ch === '—' || ch === '\n') break;
    if (/\s/.test(ch) && /^[a-z]/.test(text.slice(i + 1, i + 3))) break; // prose resuming
  }
  return text.slice(startAt, i).trim();
}

/** Extracts every `<cell> =<formula>` pair a subtask description spells out, in order. */
export function extractDescribedFormulas(description: string): DescribedFormula[] {
  const results: DescribedFormula[] = [];
  let match: RegExpExecArray | null;
  CELL_FORMULA_START.lastIndex = 0;
  while ((match = CELL_FORMULA_START.exec(description)) !== null) {
    const [, colLetter, rowText] = match;
    const row = Number(rowText) - 1;
    const col = letterToColIndex(colLetter.toUpperCase());
    if (!Number.isFinite(row) || row < 0 || !Number.isFinite(col) || col < 0) continue;

    const formulaStart = match.index + match[0].length - 1; // include the '='
    const formula = readFormula(description, formulaStart);
    if (formula === '=' || formula.length < 2) continue; // e.g. bare "A1=…" with no real formula

    results.push({
      row,
      col,
      cellRef: `${colLetter.toUpperCase()}${rowText}`,
      formula,
    });
    CELL_FORMULA_START.lastIndex = formulaStart + formula.length;
  }
  return results;
}

function normalizeFormula(formula: string): string {
  return formula.replace(/\s+/g, '').toLowerCase();
}

@Injectable()
export class FormulaConformanceChecker {
  check(states: SubtaskActionSlice[]): CheckerResult {
    const subtaskResults = states.map(({ subtask, actions }) => {
      const described = extractDescribedFormulas(subtask.description);
      if (described.length === 0) {
        return {
          subtaskId: subtask.id,
          passed: true,
          feedback: 'No formulas spelled out in the description — nothing to conform to',
          issues: [],
        };
      }

      const writtenByCell = new Map<string, string>();
      for (const action of actions) {
        if (typeof action.formula !== 'string' || !action.formula.startsWith('=')) continue;
        if (typeof action.row !== 'number' || typeof action.col !== 'number') continue;
        writtenByCell.set(`${action.row}:${action.col}`, action.formula);
      }

      const failures: string[] = [];
      for (const expected of described) {
        const written = writtenByCell.get(`${expected.row}:${expected.col}`);
        // Not written at all is CompletenessChecker/ComputedColumnChecker's
        // concern, not this checker's — only flag a formula that WAS written
        // but does not match.
        if (written === undefined) continue;
        if (normalizeFormula(written) !== normalizeFormula(expected.formula)) {
          failures.push(
            `${expected.cellRef} should be ${expected.formula} (per the subtask description) but was written as ${written}`,
          );
        }
      }

      return {
        subtaskId: subtask.id,
        passed: failures.length === 0,
        feedback:
          failures.length === 0
            ? 'Written formulas match the subtask description'
            : failures.join('; '),
        issues: failures.map((description) => ({
          severity: 'error' as const,
          subtaskId: subtask.id,
          description,
          suggestion: 'Rewrite the formula to match exactly what the subtask description specified',
        })),
      };
    });

    return buildCheckerResult(subtaskResults);
  }
}

