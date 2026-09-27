import {
  extractDescribedFormulas,
  FormulaConformanceChecker,
} from '../src/agents/checkers/formula-conformance.checker';
import { Action, SubTask } from '../src/agents/types/agent.types';
import { SubtaskActionSlice } from '../src/agents/checkers/checker.types';

const subtask = (overrides: Partial<SubTask> = {}): SubTask => ({
  id: 's1',
  description: '',
  targetSheet: 'Main',
  dependsOn: [],
  estimatedActions: 1,
  ...overrides,
});

function slice(description: string, actions: Action[], id = 's1'): SubtaskActionSlice {
  return { subtask: subtask({ id, description }), actions };
}

describe('extractDescribedFormulas', () => {
  it('extracts a single cell/formula pair', () => {
    expect(extractDescribedFormulas('B2 =SUM(B5:B16)')).toEqual([
      { row: 1, col: 1, cellRef: 'B2', formula: '=SUM(B5:B16)' },
    ]);
  });

  it('extracts multiple comma-separated pairs, matching the live planner shape', () => {
    // TASKS.md #328's exact live sentence.
    const description =
      "Write the title and KPI band on Main rows 1-2: B2 =SUM(B5:B16), D2 =SUM(C5:C16), F2 =SUM(D5:D16). Apply number formatting.";
    expect(extractDescribedFormulas(description)).toEqual([
      { row: 1, col: 1, cellRef: 'B2', formula: '=SUM(B5:B16)' },
      { row: 1, col: 3, cellRef: 'D2', formula: '=SUM(C5:C16)' },
      { row: 1, col: 5, cellRef: 'F2', formula: '=SUM(D5:D16)' },
    ]);
  });

  it('handles cross-sheet references and quoted string arguments without breaking on the comma inside them', () => {
    const description = 'C5 =SUMIF(January!I:I, "Paid", January!G:G)';
    expect(extractDescribedFormulas(description)).toEqual([
      { row: 4, col: 2, cellRef: 'C5', formula: '=SUMIF(January!I:I, "Paid", January!G:G)' },
    ]);
  });

  it('extracts multiple SUMIF pairs from a comma-separated list without truncating any of them at their internal commas', () => {
    // planner.prompt.ts's own worked example, verbatim.
    const description =
      'B5 =SUM(January!G:G), C5 =SUMIF(January!I:I, "Paid", January!G:G), D5 =SUMIF(January!I:I, "Pending", January!G:G), B6 =SUM(February!G:G)';
    expect(extractDescribedFormulas(description)).toEqual([
      { row: 4, col: 1, cellRef: 'B5', formula: '=SUM(January!G:G)' },
      { row: 4, col: 2, cellRef: 'C5', formula: '=SUMIF(January!I:I, "Paid", January!G:G)' },
      { row: 4, col: 3, cellRef: 'D5', formula: '=SUMIF(January!I:I, "Pending", January!G:G)' },
      { row: 5, col: 1, cellRef: 'B6', formula: '=SUM(February!G:G)' },
    ]);
  });

  it('reads the live #320 text without crashing or hanging, extracting nothing meaningful from its truncated "A1=…"', () => {
    // The exact live description from TASKS.md #320 — messy prose with an
    // ellipsis-truncated, contentless "A1=…" rather than a real formula.
    const description =
      "Write the title and KPI band on Main rows 1-2, with each label directly above its value in the SAME column: A1=… The KPI cells sum the Monthly Totals table's own columns — do not re-derive cross-sheet formulas here. Apply number formatting.";
    // Must not throw, and must not fabricate a formula out of an ellipsis.
    expect(() => extractDescribedFormulas(description)).not.toThrow();
    expect(extractDescribedFormulas(description).every((f) => f.formula.length > 1)).toBe(true);
  });

  it('returns nothing for a description with no spelled-out formula', () => {
    expect(extractDescribedFormulas('Format the header row bold with a fill color.')).toEqual([]);
  });

  it('does not run away past the end of the formula into a following sentence with no punctuation between them', () => {
    // No comma/period directly after the formula, just a space then prose —
    // the whitespace+lowercase-word heuristic is what has to catch this.
    const description = 'B2 =SUM(B5:B16) then format it bold';
    expect(extractDescribedFormulas(description)).toEqual([
      { row: 1, col: 1, cellRef: 'B2', formula: '=SUM(B5:B16)' },
    ]);
  });

  it('does not stop early on a lowercase sheet name inside the formula itself', () => {
    const description = 'B2 =SUM(expenses!B2:B10)';
    expect(extractDescribedFormulas(description)).toEqual([
      { row: 1, col: 1, cellRef: 'B2', formula: '=SUM(expenses!B2:B10)' },
    ]);
  });

  it('terminates at an em dash the way it does at a period', () => {
    const description = 'B2 =SUM(B5:B16) — never re-derive this elsewhere';
    expect(extractDescribedFormulas(description)).toEqual([
      { row: 1, col: 1, cellRef: 'B2', formula: '=SUM(B5:B16)' },
    ]);
  });

  it('ignores a bare cell reference with no formula', () => {
    expect(extractDescribedFormulas('Put the title in A1.')).toEqual([]);
  });
});

describe('FormulaConformanceChecker', () => {
  const checker = new FormulaConformanceChecker();

  it('passes when the written formula matches the description exactly', () => {
    const result = checker.check([
      slice('B2 =SUM(B5:B16)', [
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(B5:B16)' },
      ]),
    ]);
    expect(result.passed).toBe(true);
  });

  it('passes when the written formula differs only in whitespace/case', () => {
    const result = checker.check([
      slice('B2 =SUM(B5:B16)', [
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '= sum( B5 : B16 )' },
      ]),
    ]);
    expect(result.passed).toBe(true);
  });

  // The exact live failure: run_1790326084342_hctud2t's p3_s4.
  it('fails when the written formula deviates from what the description spelled out (TASKS.md #328)', () => {
    const result = checker.check([
      slice('B2 =SUM(B5:B16)', [
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(B5:B10)' },
      ]),
    ]);
    expect(result.passed).toBe(false);
    expect(result.issues[0].description).toContain('B2 should be =SUM(B5:B16)');
    expect(result.issues[0].description).toContain('written as =SUM(B5:B10)');
  });

  it('checks every spelled-out cell independently — one wrong formula does not hide another', () => {
    const result = checker.check([
      slice('B2 =SUM(B5:B16), D2 =SUM(C5:C16)', [
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(B5:B16)' },
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 3, formula: '=SUM(C5:C10)' },
      ]),
    ]);
    expect(result.passed).toBe(false);
    expect(result.issues).toHaveLength(1);
    expect(result.issues[0].description).toContain('D2');
  });

  it('does not flag a cell the description never named — elaboration is not deviation', () => {
    const result = checker.check([
      slice('B2 =SUM(B5:B16)', [
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(B5:B16)' },
        { type: 'SET_FORMULA', sheetName: 'Main', row: 2, col: 1, formula: '=AVERAGE(B5:B16)' },
      ]),
    ]);
    expect(result.passed).toBe(true);
  });

  it('does not flag a described cell that was never written at all — that belongs to CompletenessChecker/ComputedColumnChecker', () => {
    const result = checker.check([slice('B2 =SUM(B5:B16)', [])]);
    expect(result.passed).toBe(true);
  });

  it('passes trivially when the description has no spelled-out formula', () => {
    const result = checker.check([
      slice('Format the header row bold.', [
        { type: 'FORMAT_RANGE', sheetName: 'Main', row: 0, col: 0, rowCount: 1, colCount: 5, format: { bold: true } },
      ]),
    ]);
    expect(result.passed).toBe(true);
  });

  it('grades multiple subtasks independently', () => {
    const result = checker.check([
      slice('B2 =SUM(B5:B16)', [
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(B5:B10)' },
      ], 's1'),
      slice('D2 =SUM(C5:C16)', [
        { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 3, formula: '=SUM(C5:C16)' },
      ], 's2'),
    ]);
    expect(result.passed).toBe(false);
    expect(result.subtaskResults.find((r) => r.subtaskId === 's1')?.passed).toBe(false);
    expect(result.subtaskResults.find((r) => r.subtaskId === 's2')?.passed).toBe(true);
  });
});
