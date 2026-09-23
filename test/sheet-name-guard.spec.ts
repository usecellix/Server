import { isPlausibleSheetName } from '../src/agents/utils/sheet-name.util';
import { ensureReferencedSheetsPlanned } from '../src/agents/utils/plan-coverage.util';
import { reconcileRun } from '../src/agents/utils/reconcile.util';
import { Action, PlannerOutput, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Live incident (TASKS.md #290): a formula built its sheet reference
 * dynamically —
 *   INDIRECT("'"&TEXT(DATE(2026,ROW(INDIRECT("1:12")),1),"mmmm")&"'!A:A")
 * — and the sheet-reference regex matched the text between the apostrophes as
 * if it were a name. The coverage net then created a REAL sheet called
 * `"&TEXT(DATE(2026,ROW(INDIRECT("1:12")),1),"mmmm")&"` and put it in the
 * workbook (confirmed in run_1790104218969_tap3irw: subtask `auto_sheet_1`,
 * one ADD_SHEET action, that string as both `name` and `sheetName`).
 */

const JUNK = '"&TEXT(DATE(2026,ROW(INDIRECT("1:12")),1),"mmmm")&"';

describe('isPlausibleSheetName (TASKS.md #290)', () => {
  it('rejects the live formula fragment', () => {
    expect(isPlausibleSheetName(JUNK)).toBe(false);
  });

  it("accepts the sheet names people actually use, including ones with & and parentheses", () => {
    for (const name of ['January', 'Lists', 'Main', 'P&L', 'Q1 (Draft)', 'Jan 2026', 'FY24-25']) {
      expect(isPlausibleSheetName(name)).toBe(true);
    }
  });

  it("enforces Excel's own rules rather than guessing", () => {
    expect(isPlausibleSheetName('a'.repeat(31))).toBe(true);
    expect(isPlausibleSheetName('a'.repeat(32))).toBe(false); // over the limit
    for (const bad of ['Jan/Feb', 'Jan\\Feb', 'Q1?', 'Q1*', 'Data[1]', 'Sheet:1']) {
      expect(isPlausibleSheetName(bad)).toBe(false);
    }
    expect(isPlausibleSheetName('')).toBe(false);
    expect(isPlausibleSheetName('   ')).toBe(false);
  });
});

describe('the coverage net no longer creates a sheet from a formula fragment (TASKS.md #290)', () => {
  const context: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [{
      name: 'Sheet1', usedRange: 'A1:A1', rowCount: 1, columnCount: 1,
      values: [['']], formulas: [['']], numberFormats: [['General']],
      structure: 'data_table', headerRowIndex: 0,
    }],
    namedRanges: [],
    tables: [],
  };

  const planWith = (description: string): PlannerOutput => ({
    subtasks: [{ id: 's1', description, targetSheet: 'Main', dependsOn: [], estimatedActions: 3 }],
    clarificationsNeeded: [],
    confidence: 'high',
    reasoning: '',
  });

  it('ignores the dynamic INDIRECT reference instead of creating a junk sheet', () => {
    const { plan, added } = ensureReferencedSheetsPlanned(
      planWith(
        'On Main write B5 =SUMPRODUCT(SUMIF(INDIRECT("\'"&TEXT(DATE(2026,ROW(INDIRECT("1:12")),1),"mmmm")&"\'!A:A"),"x"))',
      ),
      context,
    );
    expect(added).toEqual([]);
    expect(plan.subtasks.map((s) => s.targetSheet)).not.toContain(JUNK);
  });

  it('still creates a genuinely-referenced sheet that nothing plans (the behaviour this net exists for)', () => {
    const { added } = ensureReferencedSheetsPlanned(
      planWith("On Main set a dropdown sourced from Lists!$B$3:$B$20"),
      context,
    );
    expect(added).toContain('Lists');
  });
});

describe('reconciliation does not report a formula fragment as a dangling sheet (TASKS.md #290)', () => {
  it('ignores the dynamic reference', () => {
    const main: SubTask = {
      id: 'p3_s1', description: 'Build Main', targetSheet: 'Main', dependsOn: [], estimatedActions: 5,
    };
    const applied = [
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      {
        type: 'BATCH_SET',
        sheetName: 'Main',
        operations: [
          { address: 'B5', formula: '=SUM(INDIRECT("\'"&TEXT(DATE(2026,ROW(INDIRECT("1:12")),1),"mmmm")&"\'!A:A"))' },
        ],
      },
    ] as unknown as Action[];

    const { gaps } = reconcileRun({ subtasks: [main], appliedActions: applied });
    expect(gaps.filter((g) => g.kind === 'dangling-reference')).toEqual([]);
  });

  it('still reports a REAL dangling reference', () => {
    const main: SubTask = {
      id: 'p3_s1', description: 'Build Main', targetSheet: 'Main', dependsOn: [], estimatedActions: 5,
    };
    const applied = [
      { type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' },
      { type: 'BATCH_SET', sheetName: 'Main', operations: [{ address: 'B5', formula: "=SUM('February'!E:E)" }] },
    ] as unknown as Action[];

    const { gaps } = reconcileRun({ subtasks: [main], appliedActions: applied });
    expect(gaps.some((g) => g.kind === 'dangling-reference' && g.detail.includes('February'))).toBe(true);
  });
});
