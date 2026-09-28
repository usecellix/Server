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
      headerRowIndex: 0,
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

  it('allows function parentheses and structured table references', () => {
    const actions: Action[] = [
      {
        type: 'SET_FORMULA',
        sheetName: 'Sheet1',
        row: 1,
        col: 2,
        formula: '=TEXT([@Date],"yyyy-mm")',
      },
    ];
    const result = validator.validatePreApply(actions, baseContext);
    expect(result.issues.filter((i) => i.code === 'SYNTAX')).toHaveLength(0);
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

  // TASKS.md #237 — a batch that writes a row AND references it in a formula
  // in the SAME batch must validate against the sheet shape AFTER its own
  // writes apply, not the sheet shape before them. Without a shadow, this
  // failed identically on every retry — unwinnable, not flaky — because the
  // check was against a sheet that could never grow no matter how many times
  // the Executor re-emitted the same already-correct actions.
  it('validates same-batch row writes and dependent formula together when given a shadow', () => {
    const headerOnly: WorkbookContext = {
      activeSheetName: 'Sheet1',
      sheets: [
        {
          name: 'Sheet1',
          usedRange: 'A1:D1',
          rowCount: 1,
          columnCount: 4,
          values: [['Item', 'Quantity', 'Price', 'Total']],
          formulas: [['', '', '', '']],
          numberFormats: [['General', 'General', 'General', 'General']],
          structure: 'data_table',
          headerRowIndex: 0,
        },
      ],
      namedRanges: [],
      tables: [],
    };
    const actions: Action[] = [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 0, value: 'Widget' },
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: 5 },
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 2, value: 10 },
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 3, formula: '=B2*C2' },
    ];

    // Without a shadow: fails, because B2/C2 don't exist in the PRE-batch
    // 1-row context — this is the bug, reproduced.
    const withoutShadow = validator.validatePreApply(actions, headerOnly);
    expect(withoutShadow.passed).toBe(false);
    expect(withoutShadow.issues.some((i) => i.code === 'REFERENCE')).toBe(true);

    // With a shadow: passes, because B2/C2 are simulated as existing once
    // this batch's own earlier SET_CELL actions are accounted for — the fix.
    const shadow = buildShadowWorkbook(headerOnly);
    const withShadow = validator.validatePreApply(actions, headerOnly, undefined, shadow);
    expect(withShadow.passed).toBe(true);
    expect(withShadow.issues.filter((i) => i.severity === 'error')).toHaveLength(0);
  });

  it('rejects hardcoded numeric literals where formulas are expected', () => {
    const actions: Action[] = [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 2, value: 180 },
    ];
    const result = validator.checkNoHardcodedLiterals(actions);
    expect(result.passed).toBe(false);
    expect(result.reason).toMatch(/numeric literal 180/i);
  });

  it('allows formula actions for GST-style requests', () => {
    const actions: Action[] = [
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=D2*0.18' },
    ];
    const result = validator.checkNoHardcodedLiterals(actions);
    expect(result.passed).toBe(true);
  });

  it('rejects domain-tool numeric results written as SET_CELL literals', () => {
    // Spec 06: Executor must write formulas referencing DomainToolResult.data —
    // never paste computed tax/ITC amounts as hard-coded cell values.
    const domainToolItcClaimable = 1800;
    const actions: Action[] = [
      {
        type: 'SET_CELL',
        sheetName: 'ITC',
        row: 1,
        col: 3,
        value: domainToolItcClaimable,
      },
    ];
    const result = validator.checkNoHardcodedLiterals(actions);
    expect(result.passed).toBe(false);
    expect(result.reason).toMatch(/numeric literal 1800/i);
  });
});

/**
 * TASKS.md #296 — a subtask's formulas must be able to reference the sheet
 * that same subtask creates.
 *
 * Live shape, from the smoke run that followed #294: the Main subtask emitted
 * `ADD_SHEET Main` and `=SUM(B5:B16)` in ONE batch, and pre-apply validation
 * rejected every formula with `Range B5:B16 points to unknown sheet "Main"`.
 * `virtualApply` had created Main in the shadow correctly; `shadowAsContext`
 * then dropped it, because it only walked the PRE-batch sheet list. Two
 * retries later the subtask failed, its dependents were gated off, and the
 * run ended at wave 4 of 6 with Main never built.
 *
 * The mirror of #294, and the reason the negative cases below matter at
 * least as much as the positive one: #294's first fix was too blunt and
 * broke a legitimate case. A fix that makes every unknown sheet resolvable
 * would "pass" this file's first test while destroying the check itself.
 */
describe('FormulaValidatorService — sheets created inside the same batch (TASKS.md #296)', () => {
  const validator = new FormulaValidatorService();

  /** The live batch, reduced to what the bug needs: a create plus its own formulas. */
  const mainSubtaskActions: Action[] = [
    { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
    { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'Payments Dashboard' },
    { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: '=SUM(B5:B16)' },
    { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 3, formula: '=SUM(C5:C16)' },
    { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 5, formula: '=SUM(D5:D16)' },
  ];

  it('accepts a formula on a sheet the same batch creates', () => {
    const shadow = buildShadowWorkbook(baseContext);
    const result = validator.validatePreApply(mainSubtaskActions, baseContext, 'Main', shadow);

    const unknownSheet = result.issues.filter(
      (i) => i.severity === 'error' && i.message.includes('unknown sheet'),
    );
    expect(unknownSheet).toHaveLength(0);
    expect(result.passed).toBe(true);
  });

  it('resolves a sheet the batch merely writes to — that gap belongs to the #294 checker', () => {
    // Dropping the ADD_SHEET does NOT make this fail, and that is correct
    // behaviour for this layer: `virtualApply`'s `ensureSheet` conjures a
    // shadow sheet on any write (virtualApply.ts:186), so Main resolves here
    // whether or not anything created it. Catching "written to but never
    // created" is StructuralIntentChecker's job and it does so against the
    // PRE-RUN context precisely so this conjuring cannot fool it (TASKS.md
    // #294). Asserting it here instead would duplicate that check in the one
    // place structurally unable to perform it.
    const withoutCreate = mainSubtaskActions.filter((a) => a.type !== 'ADD_SHEET');
    const shadow = buildShadowWorkbook(baseContext);
    const result = validator.validatePreApply(withoutCreate, baseContext, 'Main', shadow);

    expect(
      result.issues.some((i) => i.message.includes('unknown sheet "Main"')),
    ).toBe(false);
  });

  it('still rejects a cross-sheet reference to a sheet no one creates', () => {
    const actions: Action[] = [
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      // Main exists now; Quarter4 does not, and nothing in the batch makes it.
      { type: 'SET_FORMULA', sheetName: 'Main', row: 1, col: 1, formula: "=SUM(Quarter4!B5:B16)" },
    ];
    const shadow = buildShadowWorkbook(baseContext);
    const result = validator.validatePreApply(actions, baseContext, 'Main', shadow);

    expect(result.passed).toBe(false);
    expect(
      result.issues.some((i) => i.message.includes('unknown sheet "Quarter4"')),
    ).toBe(true);
  });

  it('leaves an already-existing sheet with its real content (TASKS.md #237 unchanged)', () => {
    // The pre-existing sheet must keep its values and merely grow its bounds —
    // appending created sheets must not shadow-replace the ones already there.
    const actions: Action[] = [
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 9, col: 1, formula: '=SUM(B1:B3)' },
    ];
    const shadow = buildShadowWorkbook(baseContext);
    const result = validator.validatePreApply(actions, baseContext, 'Sheet1', shadow);

    expect(result.passed).toBe(true);
    expect(baseContext.sheets[0].values[0]).toEqual(['Name', 'Qty', 'Price']);
  });

  it('accepts a formula that reads a DIFFERENT sheet created in the same batch', () => {
    // The consolidation shape: Main sums across month sheets the same wave
    // creates. This is what the live run was ultimately trying to build.
    const actions: Action[] = [
      { type: 'ADD_SHEET', name: 'January', sheetName: 'January' },
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      { type: 'SET_FORMULA', sheetName: 'Main', row: 4, col: 1, formula: "=SUM(January!G2:G100)" },
    ];
    const shadow = buildShadowWorkbook(baseContext);
    const result = validator.validatePreApply(actions, baseContext, 'Main', shadow);

    expect(result.issues.filter((i) => i.message.includes('unknown sheet'))).toHaveLength(0);
    expect(result.passed).toBe(true);
  });
});
