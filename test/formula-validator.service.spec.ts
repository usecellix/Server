import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import { WorkbookContext } from '../src/agents/types/agent.types';
import { Action } from '../src/agents/types/agent.types';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';

const baseContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:C3',
      rowCount: 3,
      columnCount: 3,
      values: [
        ['Name', 'Qty', 'Price'],
        ['Apple', 10, 1.5],
        ['Banana', 5, 0.75],
      ],
      formulas: [['', '', ''], ['', '', ''], ['', '', '']],
      numberFormats: [['General', 'General', 'General']],
      structure: 'data_table',
    },
  ],
  namedRanges: [{ name: 'TaxRate', formula: '=0.1' }],
  tables: [],
};

describe('FormulaValidatorService', () => {
  const validator = new FormulaValidatorService();

  it('passes valid formula references', () => {
    const actions: Action[] = [
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=B2*C2' },
    ];
    const result = validator.validatePreApply(actions, baseContext);
    expect(result.passed).toBe(true);
    expect(result.issues.filter((i) => i.severity === 'error')).toHaveLength(0);
  });

  it('rejects unbalanced parentheses', () => {
    const actions: Action[] = [
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=SUM(B2:C2' },
    ];
    const result = validator.validatePreApply(actions, baseContext);
    expect(result.passed).toBe(false);
    expect(result.issues.some((i) => i.code === 'SYNTAX')).toBe(true);
  });

  it('rejects out-of-bounds cell references', () => {
    const actions: Action[] = [
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=Z99*2' },
    ];
    const result = validator.validatePreApply(actions, baseContext);
    expect(result.passed).toBe(false);
    expect(result.issues.some((i) => i.code === 'REFERENCE')).toBe(true);
  });

  it('warns on unknown named range identifiers', () => {
    const actions: Action[] = [
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=B2*UnknownRate' },
    ];
    const result = validator.validatePreApply(actions, baseContext);
    expect(result.issues.some((i) => i.code === 'NAMED_RANGE')).toBe(true);
  });

  it('detects post-apply Excel error strings in shadow cells', () => {
    const shadow = buildShadowWorkbook(baseContext);
    const actions: Action[] = [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 2, value: '#REF!' },
    ];
    const after = virtualApply(shadow, actions);
    const result = validator.checkPostApply(after, actions, baseContext);
    expect(result.passed).toBe(false);
    expect(result.issues.some((i) => i.code === 'POST_EXEC')).toBe(true);
  });

  it('extracts formulas from ADD_ROW data arrays', () => {
    const actions: Action[] = [
      { type: 'ADD_ROW', sheetName: 'Sheet1', data: ['Total', '=B2*C2', ''] },
    ];
    const result = validator.validatePreApply(actions, baseContext);
    expect(result.passed).toBe(true);
  });
});
