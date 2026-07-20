import {
  checkFormulaActionAgainstRule,
  deriveSemanticRuleFromPrompt,
  SemanticFormulaChecker,
} from '../src/agents/checkers/semantic-formula.checker';
import { WorkbookContext } from '../src/agents/types/agent.types';

const context: WorkbookContext = {
  activeSheetName: 'Purchases',
  sheets: [
    {
      name: 'Purchases',
      usedRange: 'A1:D3',
      rowCount: 3,
      columnCount: 4,
      values: [
        ['Qty', 'Unit Price', 'Tax Amount', 'Total'],
        [10, 100, '', ''],
        [2, 50, '', ''],
      ],
      formulas: [[], [], []],
      numberFormats: [[], [], []],
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('semantic formula checker', () => {
  it('derives 18% tax rule from an explicit prompt', () => {
    const rule = deriveSemanticRuleFromPrompt(
      'Set Tax Amount = 18% of Qty × Unit Price for each row',
    );
    expect(rule?.rate).toBe(0.18);
    expect(rule?.requiredColumns.length).toBeGreaterThan(0);
  });

  it('passes a correct 18% formula referencing Qty and Unit Price', () => {
    const rule = deriveSemanticRuleFromPrompt(
      'Tax Amount should be 18% of Qty × Unit Price',
    )!;
    const result = checkFormulaActionAgainstRule(
      { type: 'SET_FORMULA', sheetName: 'Purchases', formula: '=A2*B2*0.18', row: 1, col: 2 },
      rule,
      context,
    );
    expect(result.passed).toBe(true);
  });

  it('fails a wrong tax rate (fixture #7/#8 style)', () => {
    const rule = deriveSemanticRuleFromPrompt(
      'Tax Amount = 18% of Qty × Unit Price',
    )!;
    const result = checkFormulaActionAgainstRule(
      { type: 'SET_FORMULA', sheetName: 'Purchases', formula: '=A2*B2*0.12', row: 1, col: 2 },
      rule,
      context,
    );
    expect(result.passed).toBe(false);
    expect(result.reason).toMatch(/rate/i);
  });

  it('fails when formula ignores required columns', () => {
    const rule = deriveSemanticRuleFromPrompt(
      'Tax Amount = 18% of Qty × Unit Price',
    )!;
    const result = checkFormulaActionAgainstRule(
      { type: 'SET_FORMULA', sheetName: 'Purchases', formula: '=D2*0.18', row: 1, col: 2 },
      rule,
      context,
    );
    expect(result.passed).toBe(false);
  });

  it('flags wrong formulas via the checker service', () => {
    const checker = new SemanticFormulaChecker();
    const result = checker.check(
      'Add Tax Amount as 18% of Qty × Unit Price',
      [{ id: 's1', description: 'Add tax formula', targetSheet: 'Purchases', dependsOn: [], estimatedActions: 1 }],
      [
        {
          subtask: {
            id: 's1',
            description: 'Add tax formula',
            targetSheet: 'Purchases',
            dependsOn: [],
            estimatedActions: 1,
          },
          actions: [
            { type: 'SET_FORMULA', sheetName: 'Purchases', formula: '=C2*0.05', row: 1, col: 2 },
          ],
        },
      ],
      context,
    );
    expect(result.passed).toBe(false);
  });
});
