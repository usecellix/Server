import { applyPresentationPass, detectHeaderRuns } from '../src/excel-ai/utils/presentation-pass.util';
import { applyConsolidationPass, planConsolidation } from '../src/excel-ai/utils/consolidation-pass.util';
import { sheetsCreatedInBatch } from '../src/excel-ai/utils/sheet-header-state.util';
import { summarizePlanIntent } from '../src/excel-ai/utils/user-facing-response.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #158 — these passes must work for ANY prompt, not the one they were
 * built against.
 *
 * Two prompt-shaped assumptions were found and removed:
 *   - consolidation required the origin column to be the literal word "month",
 *     silently limiting the feature to calendar workbooks;
 *   - the number-format word lists were hospitality-flavoured, so payroll,
 *     invoicing and tax workbooks got nothing.
 *
 * Every case below is a DIFFERENT domain from the monthly-booking prompt.
 */

function build(
  sheets: string[],
  schema: string[],
  mainSchema: string[],
  headerRow = 17,
): SheetActionPayload[] {
  const a: SheetActionPayload[] = [];
  for (const s of sheets) a.push({ type: 'ADD_SHEET', name: s, sheetName: s });
  a.push({ type: 'ADD_SHEET', name: 'Main', sheetName: 'Main' });
  for (const s of sheets) {
    schema.forEach((h, c) => a.push({ type: 'SET_CELL', sheetName: s, row: 0, col: c, value: h }));
  }
  mainSchema.forEach((h, c) =>
    a.push({ type: 'SET_CELL', sheetName: 'Main', row: headerRow, col: c, value: h }),
  );
  return a;
}

describe('generality — consolidation is structural, not calendar-specific', () => {
  it('consolidates one sheet per REGION', () => {
    const plan = planConsolidation(
      build(
        ['North', 'South', 'East', 'West'],
        ['Rep', 'Target', 'Actual'],
        ['Region', 'Rep', 'Target', 'Actual'],
      ),
    );
    expect(plan).not.toBeNull();
    expect(plan!.sourceSheets).toEqual(['North', 'South', 'East', 'West']);
  });

  it('consolidates one sheet per DEPARTMENT', () => {
    const plan = planConsolidation(
      build(
        ['Engineering', 'Sales', 'Support'],
        ['Employee', 'Salary', 'Bonus'],
        ['Department', 'Employee', 'Salary', 'Bonus'],
      ),
    );
    expect(plan).not.toBeNull();
    expect(plan!.sourceSheets).toHaveLength(3);
  });

  it('consolidates one sheet per QUARTER, quoting names with spaces', () => {
    const actions = build(
      ['Q1 2026', 'Q2 2026', 'Q3 2026'],
      ['Invoice No', 'Client', 'Amount'],
      ['Quarter', 'Invoice No', 'Client', 'Amount'],
    );
    expect(planConsolidation(actions)).not.toBeNull();
    const out = applyConsolidationPass(actions, { dynamicArrays: true });
    expect(out[out.length - 1].formula).toContain("'Q1 2026'!");
  });

  it('still refuses when the schemas do not actually match', () => {
    // Main's own rollup: starts with a key column, but its data columns are a
    // summary rather than the sibling schema. This is what exact-match is for.
    const plan = planConsolidation(
      build(['North', 'South'], ['Rep', 'Target', 'Actual'], ['Region', 'Total Target', 'Total Actual']),
    );
    expect(plan).toBeNull();
  });

  it('still needs at least two sibling sheets', () => {
    expect(
      planConsolidation(build(['North'], ['Rep', 'Target', 'Actual'], ['Region', 'Rep', 'Target', 'Actual'])),
    ).toBeNull();
  });
});

describe('generality — number formats span business domains', () => {
  function formatFor(header: string): string | undefined {
    const actions = build(['A', 'B'], ['Key', header, 'Other'], []);
    const out = applyPresentationPass(actions, {});
    return out.find((x) => x.type === 'FORMAT_RANGE' && x.row === 1 && x.col === 1)?.format
      ?.numberFormat;
  }

  it.each([
    ['Salary'],
    ['Commission'],
    ['Bonus'],
    ['Reimbursement'],
    ['Invoice Amount'],
    ['Discount'],
    ['Refund'],
    ['Deposit'],
    ['GST'],
    ['TDS'],
    ['VAT'],
    ['Freight'],
    ['Premium'],
    ['Interest'],
    ['Revenue'],
    ['Expense'],
    ['Budget'],
    ['Margin'],
    ['Profit'],
  ])('formats %s as currency', (header) => {
    expect(formatFor(header)).toBe('#,##0.00');
  });

  it.each([['Quantity'], ['Units'], ['Headcount'], ['Rooms'], ['Tickets'], ['Attendance']])(
    'formats %s as a whole-number count',
    (header) => {
      expect(formatFor(header)).toBe('#,##0');
    },
  );

  it.each([
    ['Guest Name'],
    ['Payment Status'],
    ['Unit No'],
    ['Invoice Number'],
    ['Account Code'],
    ['Supplier Name'],
    ['GSTIN'],
    ['Payment Mode'],
  ])('never numeric-formats %s (text or identifier)', (header) => {
    expect(formatFor(header)).toBeUndefined();
  });
});

describe('generality — intent bullets group any repeated shape', () => {
  it.each([
    [['North', 'South', 'East', 'West']],
    [['Alpha Corp', 'Beta Ltd', 'Gamma Inc']],
    [['Q1', 'Q2', 'Q3', 'Q4']],
  ])('collapses %s into one counted line', (sheets) => {
    const bullets = summarizePlanIntent(
      sheets.map((s, i) => ({
        id: `s${i}`,
        description: `Create sheet '${s}' and set A1:C1 headers [Item, Qty, Amount]`,
        targetSheet: s,
      })),
    );
    expect(bullets).toHaveLength(1);
    expect(bullets[0]).toContain(`${sheets.length} sheets`);
  });
});

describe('generality — styling applies to any created sheet', () => {
  it('styles a non-calendar workbook identically', () => {
    const actions = build(['Engineering', 'Sales'], ['Employee', 'Salary', 'Bonus'], []);
    const out = applyPresentationPass(actions, {});
    for (const sheet of ['Engineering', 'Sales']) {
      expect(
        out.some(
          (a) =>
            a.type === 'FORMAT_RANGE' &&
            a.sheetName === sheet &&
            a.format?.fillColor === '#2F5597',
        ),
      ).toBe(true);
      expect(out.some((a) => a.type === 'FREEZE_PANES' && a.sheetName === sheet)).toBe(true);
      expect(out.some((a) => a.type === 'AUTOFIT_COLUMNS' && a.sheetName === sheet)).toBe(true);
    }
  });

  it('leaves sheets the batch did not create completely alone', () => {
    const actions: SheetActionPayload[] = [
      { type: 'SET_CELL', sheetName: 'ExistingData', row: 0, col: 0, value: 'Employee' },
      { type: 'SET_CELL', sheetName: 'ExistingData', row: 0, col: 1, value: 'Salary' },
      { type: 'SET_CELL', sheetName: 'ExistingData', row: 0, col: 2, value: 'Bonus' },
    ];
    expect(detectHeaderRuns(actions, sheetsCreatedInBatch(actions))).toHaveLength(0);
    expect(applyPresentationPass(actions, {})).toEqual(actions);
  });
});
