import {
  changeCase,
  changeCaseLane,
  classifyRows,
  dynamicListLane,
  fillBlanksLane,
  findHeaderRow,
  groupByLane,
  indianFormatFor,
  lakhCroreLane,
  removeRowTypesLane,
  subtotalsLane,
  TEMPLATES,
  templatesLane,
} from './stage3.lanes';
import { buildCalendar, complianceCalendarLane, entityFlags, validationPresetLane } from './stage4.lanes';
import { getRuleTable } from '../../rules/rule-tables';
import { matchingLanes, runDeterministicLanes } from './lane-registry';
import { cellsOf, makeCtx, serial } from './lane-test-utils';
import { LaneResult } from './lane.types';

const answerOf = (r: unknown) => (r as LaneResult).answer;
const actionsOf = (r: unknown) => ((r as { actions?: Array<Record<string, unknown>> }).actions ?? []);

describe('P3.1 change case', () => {
  it.each([
    ['abc pvt ltd', 'Abc Pvt Ltd'],
    ['XYZ TRADERS LLP', 'Xyz Traders LLP'],
    ['mehta and sons huf', 'Mehta And Sons HUF'],
    ['RAJ & CO', 'Raj & Co'],
  ])('%s → %s', (a, b) => expect(changeCase(a, 'proper')).toBe(b));

  it('skips the GSTIN column', () => {
    const vals = [['Name', 'GSTIN'], ['abc pvt ltd', '27AABCA1234F1Z9'], ['sharma enterprises', '33CCCHD9876L1ZN']];
    const res = changeCaseLane.run(makeCtx(vals, { message: 'Change the names to proper case' }));
    const cells = cellsOf(res as LaneResult);
    expect(cells.A2).toBe('Abc Pvt Ltd');
    expect(Object.keys(cells).some((a) => a.startsWith('B'))).toBe(false);
    expect(answerOf(res)).toMatch(/Not touched.*B \(GSTIN\)/);
  });
});

describe('P3.2 fill blanks', () => {
  const T_BLANKS = [
    ['Group', 'Party', 'Amount'],
    ['Sundry Debtors', 'A Co', 1000],
    [null, 'B Co', 2000],
    [null, 'C Co', 3000],
    ['Sundry Creditors', 'D Co', 4000],
    [null, 'E Co', 5000],
    [' ', 'F Co', 600],
  ];
  it('fills only truly blank cells and reports the space-only cell', () => {
    const res = fillBlanksLane.run(makeCtx(T_BLANKS, { message: 'Fill the blank cells in column A with the value above' }));
    expect(cellsOf(res as LaneResult)).toEqual({ A3: 'Sundry Debtors', A4: 'Sundry Debtors', A6: 'Sundry Creditors' });
    expect(answerOf(res)).toMatch(/Left alone: A7/);
  });
  it('treats spaces as blank when asked', () => {
    const res = fillBlanksLane.run(makeCtx(T_BLANKS, { message: 'Fill the blank cells in column A with the value above and treat spaces as blank' }));
    expect(cellsOf(res as LaneResult).A7).toBe('Sundry Creditors');
  });
  it('fills with zero', () => {
    const res = fillBlanksLane.run(makeCtx([['A', 'B'], [1, 'x'], [null, 'y']], { message: 'fill blanks in column A with zero' }));
    expect(cellsOf(res as LaneResult)).toEqual({ A3: 0 });
  });
});

describe('P3.3 remove row types', () => {
  const T_ROWTYPES = [
    ['XYZ Traders - Ledger Report', null, null],
    [null, null, null],
    ['Date', 'Particulars', 'Amount'],
    [serial(2026, 4, 1), 'Sale 1', 1000],
    [serial(2026, 4, 2), 'Sale 2', 2000],
    [serial(2026, 4, 3), 'Sale 3', 3000],
    [serial(2026, 4, 4), 'Sale 4', 4000],
    [null, 'Sub Total', 10000],
    [null, null, null],
    ['Date', 'Particulars', 'Amount'],
    [serial(2026, 4, 5), 'Sale 5', 500],
    [serial(2026, 4, 6), 'Sale 6', 600],
    [serial(2026, 4, 7), 'Sale 7', 700],
    [serial(2026, 4, 8), 'Sale 8', 800],
    [null, 'Sub Total', 2600],
    [null, 'Grand Total', 12600],
  ];
  it('classifies every row', () => {
    const h = findHeaderRow(T_ROWTYPES);
    expect(h).toBe(2);
    const t = classifyRows(T_ROWTYPES, h);
    expect(t.filter((x) => x === 'data')).toHaveLength(8);
    expect(t.filter((x) => x === 'subtotal')).toHaveLength(2);
    expect(t.filter((x) => x === 'grand-total')).toHaveLength(1);
    expect(t.filter((x) => x === 'header-repeat')).toHaveLength(1);
    expect(t[0]).toBe('title');
  });
  it('writes the 8 data rows to a NEW sheet with a tie-out line', () => {
    const res = removeRowTypesLane.run(makeCtx(T_ROWTYPES, { sheetName: 'T_RowTypes', message: 'Remove the blank, repeated-header, subtotal and total rows from this report' }));
    const actions = actionsOf(res);
    expect(actions[0]).toMatchObject({ type: 'CREATE_SHEET', uniqueName: true, name: 'T_RowTypes - Clean' });
    expect((actions[1].rows as unknown[][]).length).toBe(8);
    expect(answerOf(res)).toMatch(/data sum 12,600 = Grand Total 12,600/);
    expect(actions.some((a) => a.sheetName === 'T_RowTypes' && a.type !== 'CREATE_SHEET')).toBe(false);
  });
  it('keeps the title when asked', () => {
    const res = removeRowTypesLane.run(makeCtx(T_ROWTYPES, { sheetName: 'T_RowTypes', message: 'Remove the subtotal and total rows but keep the title' }));
    expect(actionsOf(res)[1]).toMatchObject({ headers: ['XYZ Traders - Ledger Report', '', ''] });
  });
});

describe('P3.4 lakh / crore', () => {
  it.each([
    [1234567, '##\\,##\\,##0'],
    [12345678, '##\\,##\\,##\\,##0'],
    [99999, '#,##0'],
    [150000000.5, '##\\,##\\,##\\,##0.00'],
    [-2500000, '##\\,##\\,##0'],
  ])('%d → %s', (n, f) => expect(indianFormatFor(n)).toBe(f));

  it('changes formats only, never values', () => {
    const vals = [['Amount'], [1234567], [12345678], [99999], [150000000.5], [-2500000]];
    const res = lakhCroreLane.run(makeCtx(vals, { message: 'Format column A in Indian lakh and crore style' }));
    const actions = actionsOf(res);
    expect(actions.every((a) => a.type === 'FORMAT_RANGE')).toBe(true);
    expect(answerOf(res)).toMatch(/12,34,567 · 1,23,45,678 · 99,999 · 15,00,00,000.50 · -25,00,000/);
  });

  it('lakhs view uses a helper column of formulas', () => {
    const vals = [['Amount'], [1234567]];
    const res = lakhCroreLane.run(makeCtx(vals, { message: 'show column A in lakhs' }));
    expect(cellsOf(res as LaneResult)).toMatchObject({ B1: 'Amount (₹ lakhs)', B2: '=A2/100000' });
    expect(answerOf(res)).toMatch(/12\.35/);
  });
});

describe('P3.16 unique sorted list', () => {
  const vals = [['Vendor'], ['Zeta'], ['ABC'], ['abc'], ['Mid']];
  it('previews the spill range and writes SORT(UNIQUE())', () => {
    const res = dynamicListLane.run(makeCtx(vals, { message: 'Give me a unique sorted list of vendors in D2' }));
    expect(cellsOf(res as LaneResult)).toEqual({ D2: '=SORT(UNIQUE(A2:A5))' });
    expect(answerOf(res)).toMatch(/spills into \*\*D2:D4\*\*/);
  });
  it('warns about a collision before Accept', () => {
    const blocked = vals.map((r, i) => (i === 2 ? [...r, '', '', 'X'] : r));
    const res = dynamicListLane.run(makeCtx(blocked, { message: 'Give me a unique sorted list of vendors in D2' }));
    expect((res as LaneResult).kind).toBe('answer');
    expect(answerOf(res)).toMatch(/D3 is not empty.*#SPILL!/);
  });
  it('falls back to static values without dynamic arrays', () => {
    const res = dynamicListLane.run(makeCtx(vals, { message: 'Give me a unique sorted list of vendors in D2', dynamicArrays: false }));
    expect(cellsOf(res as LaneResult)).toEqual({ D2: 'ABC', D3: 'Mid', D4: 'Zeta' });
  });
});

describe('P3.18 group-by', () => {
  const T_DAYS = [
    ['Invoice', 'Party', 'Invoice Date', 'Amount', 'Credit Days'],
    ['INV-001', 'ABC Traders', 1, 10000, 30],
    ['INV-002', 'XYZ Pvt Ltd', 1, 25000, 45],
    ['INV-003', 'Mehta & Sons', 1, 5000, 0],
    ['INV-004', 'Kumar Agency', 1, 7000, 15],
    ['INV-005', 'Shah Bros', 1, 3000, 30],
  ];
  it('totals tie to the source (50,000)', () => {
    const res = groupByLane.run(makeCtx(T_DAYS, { sheetName: 'T_Days', message: 'Total Amount by Party and Credit Days' }));
    expect(answerOf(res)).toMatch(/groups total 50,000 = source total 50,000 ✓/);
  });
  it('as SUMIFS formulas', () => {
    const res = groupByLane.run(makeCtx(T_DAYS, { sheetName: 'T_Days', message: 'Total Amount by Party and Credit Days as SUMIFS formulas' }));
    const batch = actionsOf(res).find((a) => a.type === 'BATCH_SET') as { operations: Array<{ formula?: string }> };
    expect(batch.operations[0].formula).toMatch(/^=SUMIFS\('T_Days'!\$D\$2:\$D\$6,'T_Days'!\$B\$2:\$B\$6,A2,'T_Days'!\$E\$2:\$E\$6,B2\)$/);
  });
});

describe('P3.19 subtotals on a copy', () => {
  const FILLED = [
    ['Group', 'Party', 'Amount'],
    ['Sundry Debtors', 'A Co', 1000],
    ['Sundry Debtors', 'B Co', 2000],
    ['Sundry Debtors', 'C Co', 3000],
    ['Sundry Creditors', 'D Co', 4000],
    ['Sundry Creditors', 'E Co', 5000],
    ['Sundry Creditors', 'F Co', 600],
  ];
  it('6,000 / 9,600 / 15,600 on a new sheet', () => {
    const res = subtotalsLane.run(makeCtx(FILLED, { sheetName: 'T_Blanks', message: 'Add subtotals per Group on a copy' }));
    const a = answerOf(res);
    expect(a).toMatch(/Sundry Debtors: 6,000/);
    expect(a).toMatch(/Sundry Creditors: 9,600/);
    expect(a).toMatch(/Grand total: 15,600/);
    expect(actionsOf(res).every((x) => x.sheetName !== 'T_Blanks')).toBe(true);
    const batch = actionsOf(res).find((x) => x.type === 'BATCH_SET') as { operations: Array<{ address: string; formula: string }> };
    expect(batch.operations.map((o) => o.formula)).toEqual(['=SUBTOTAL(9,C2:C4)', '=SUBTOTAL(9,C6:C8)', '=SUBTOTAL(9,C2:C9)']);
  });
});

describe('P3.23 / P3.24 / P4.7 templates', () => {
  it('lists templates', () => {
    const res = templatesLane.run(makeCtx([[]], { message: 'Show templates' }));
    expect((res as LaneResult).kind).toBe('answer');
    expect(TEMPLATES.map((t) => t.id)).toEqual(expect.arrayContaining(['pbc', 'notice', 'gst-recon', 'tds', 'far', 'ageing', 'compliance-tracker']));
  });
  it.each([
    ['Create a document request tracker', ['Item', 'Requested on', 'Due', 'Received', 'Status', 'Remarks']],
    ['Create a notice tracker', ['Authority', 'Notice type', 'DIN', 'Date', 'Reply due', 'Status', 'Remarks']],
    ['Create an ageing template', ['Party', 'Invoice No', 'Invoice Date', 'Amount', 'Days Outstanding', '0-30', '31-60', '61-90', '91-180', '>180']],
  ])('%s', (msg, headers) => {
    const res = templatesLane.run(makeCtx([[]], { message: msg }));
    const acts = actionsOf(res);
    expect(acts[1]).toMatchObject({ type: 'WRITE_TABLE', headers, rows: [] });
  });
  it('trackers get a status dropdown', () => {
    const acts = actionsOf(templatesLane.run(makeCtx([[]], { message: 'Create a notice tracker' })));
    expect(acts.find((a) => a.type === 'DATA_VALIDATION')).toMatchObject({ range: 'F2:F1000' });
  });
});

describe('P4.1 compliance calendar', () => {
  const rules = getRuleTable('compliance-calendar').data;
  it('monthly GST + TDS company: GSTR-1 11th, 3B 20th, TDS 7th with March → 30 Apr, advance tax 4 dates', () => {
    const rows = buildCalendar(rules, entityFlags('company with monthly GST and TDS'), 2026);
    const d = (ret: string, period: string) => {
      const r = rows.find((x) => x.ret === ret && x.period.startsWith(period));
      return r ? new Date(r.due).toISOString().slice(0, 10) : null;
    };
    expect(d('GSTR-1', 'Apr')).toBe('2026-05-11');
    expect(d('GSTR-3B', 'Dec')).toBe('2027-01-20');
    expect(d('TDS payment', 'Apr')).toBe('2026-05-07');
    expect(d('TDS payment', 'Mar')).toBe('2027-04-30');
    expect(rows.filter((r) => r.ret === 'Advance tax').map((r) => new Date(r.due).toISOString().slice(0, 10))).toEqual([
      '2026-06-15', '2026-09-15', '2026-12-15', '2027-03-15',
    ]);
    expect(rows.filter((r) => r.ret === 'GSTR-1')).toHaveLength(12);
  });
  it('lane cites the rule source, adds a status dropdown and the 7-day highlight', () => {
    const res = complianceCalendarLane.run(makeCtx([[]], { message: 'Create a compliance calendar for FY 2026-27 for a company with monthly GST and TDS' }));
    expect(answerOf(res)).toMatch(/Rule source:.*compliance-calendar v2026-27\.v1 \(CA to verify\)/);
    const acts = actionsOf(res);
    expect(acts.some((a) => a.type === 'DATA_VALIDATION')).toBe(true);
    expect(acts.some((a) => a.type === 'CONDITIONAL_FORMAT' && /TODAY\(\)<=7/.test(JSON.stringify(a.rule)))).toBe(true);
    expect((acts[1].rows as unknown[][])[0][8]).toMatch(/compliance-calendar/);
  });
});

describe('P4.2 validation presets', () => {
  it('GST rates come from the rule table', () => {
    const res = validationPresetLane.run(makeCtx([['A', 'B', 'C', 'Rate'], [1, 2, 3, 18]], { message: 'Add a dropdown of GST rates to column D' }));
    expect(actionsOf(res)[0]).toMatchObject({
      type: 'DATA_VALIDATION',
      validation: { kind: 'list', listSource: ['0', '0.1', '0.25', '3', '5', '12', '18', '28'] },
    });
    expect(answerOf(res)).toMatch(/validation-presets v/);
  });
});

describe('lane registry', () => {
  it('never offers a write outside Action mode', () => {
    expect(matchingLanes('Fill the blank cells in column A with the value above', false)).toHaveLength(0);
  });
  it('does not claim ordinary prompts', () => {
    for (const m of ['Sort by Amount descending', 'Create a bar chart of Amount by Party', 'Put the number 1 in cells A1 to A5', 'Explain this formula', 'Highlight amounts above 20000']) {
      expect(matchingLanes(m, true).map((l) => l.id)).toEqual([]);
    }
  });
  it('renames an output sheet that already exists', async () => {
    const ctx = makeCtx([[]], { message: 'Create a fixed asset register', sheetNames: ['Fixed Asset Register'] });
    const out = await runDeterministicLanes(ctx);
    expect(out?.result.kind).toBe('actions');
    expect((out!.result as { actions: Array<{ name?: string }> }).actions[0].name).toBe('Fixed Asset Register (2)');
  });
});

import { gstr3bTableWiseLane, parseTable4Sheet } from './recon.lanes';

describe('P3.8 lane: table-wise 3B vs 2B from the active sheet', () => {
  const SHEET = [
    ['Table', '3B IGST', '3B CGST', '3B SGST', '2B IGST', '2B CGST', '2B SGST'],
    ['4A(5) All other ITC', 100000, 50000, 50000, 95000, 50000, 50000],
    ['4A(3) Inward RCM', 18000, 0, 0, 0, 0, 0],
    ['4B(1) Reversals', 2000, 0, 0, 0, 0, 0],
  ];
  it('parses the layout', () => expect(parseTable4Sheet(SHEET)).toHaveLength(3));
  it('answers with per-table variance and tie-out, report goes to a new sheet', () => {
    expect(gstr3bTableWiseLane.detect('Compare GSTR-3B with GSTR-2B table-wise')).toBe(true);
    const res = gstr3bTableWiseLane.run(makeCtx(SHEET, { message: 'Compare GSTR-3B with GSTR-2B table-wise' }));
    expect(answerOf(res)).toMatch(/table variances 21,000 = total variance 21,000/);
    expect(actionsOf(res)[0]).toMatchObject({ type: 'CREATE_SHEET', uniqueName: true });
  });
  it('explains the layout when the sheet does not have it', () => {
    const res = gstr3bTableWiseLane.run(makeCtx([['A']], { message: 'Compare GSTR-3B with GSTR-2B table-wise' }));
    expect((res as LaneResult).kind).toBe('answer');
  });
});
