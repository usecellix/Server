import {
  countDataRows,
  describeEmptySheet,
} from '../src/agents/utils/data-row-count.util';
import { buildPlannerUserMessage } from '../src/agents/prompts/planner.prompt';
import { WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #84 / COMPETITIVE_STUDY_SHORTCUT.md Trial 2.
 *
 * The planner was told `rowCount`, which counts the header row and every
 * pre-provisioned template row. A month sheet of "headers + 120 seeded-formula
 * rows" therefore reported 121 rows while holding zero bookings — so a request
 * whose premise was false ("the old bank accounts are still in the sheets")
 * looked satisfiable and got planned against.
 */
describe('countDataRows', () => {
  it('counts zero for a header-only sheet', () => {
    expect(countDataRows([['Unit No', 'Guest', 'Rate']], 0)).toBe(0);
  });

  it('counts zero for the Trial 2 shape: headers + template rows whose formulas evaluate blank', () => {
    // Loaded values are COMPUTED, so `=(F-E)*G` over empty inputs arrives as ''.
    const values: unknown[][] = [['Unit No', 'Guest', 'Total Amount']];
    for (let i = 0; i < 120; i += 1) values.push(['', '', '']);
    expect(values).toHaveLength(121); // what rowCount would report
    expect(countDataRows(values, 0)).toBe(0); // what is actually there
  });

  it('counts only rows with real entered values', () => {
    const values: unknown[][] = [
      ['Unit No', 'Guest', 'Total'],
      ['101', 'Ada', 500],
      ['', '', ''],
      ['102', 'Grace', 750],
      [null, undefined, '   '],
    ];
    expect(countDataRows(values, 0)).toBe(2);
  });

  it('treats 0 and false as real data, not blanks', () => {
    expect(countDataRows([['H'], [0], [false]], 0)).toBe(2);
  });

  it('respects a non-zero header row index', () => {
    const values: unknown[][] = [
      ['Bookings Report'],
      ['Unit No', 'Guest'],
      ['101', 'Ada'],
    ];
    expect(countDataRows(values, 1)).toBe(1);
  });

  it('never throws on malformed input', () => {
    expect(countDataRows([], 0)).toBe(0);
    expect(countDataRows(undefined as unknown as unknown[][], 0)).toBe(0);
    expect(countDataRows([['H'], undefined as unknown as unknown[]], 0)).toBe(0);
  });
});

describe('describeEmptySheet', () => {
  it('stays silent for a populated sheet, so this never becomes noise', () => {
    expect(describeEmptySheet('Jan', 12, 13)).toBeNull();
  });

  it('calls out a scaffolded-but-empty sheet and explains the misleading rowCount', () => {
    const note = describeEmptySheet('Jan', 0, 121);
    expect(note).toContain('0 data rows');
    expect(note).toContain('121 rows are structure, not entered data');
  });

  it('describes a truly empty sheet plainly', () => {
    expect(describeEmptySheet('Sheet1', 0, 1)).toContain('empty sheet');
  });
});

describe('buildPlannerUserMessage — empty-sheet disclosure', () => {
  function sheet(
    name: string,
    values: unknown[][],
    rowCount: number,
  ): WorkbookContext['sheets'][number] {
    return {
      name,
      usedRange: `A1:C${rowCount}`,
      rowCount,
      columnCount: values[0]?.length ?? 0,
      values,
      formulas: values.map((r) => r.map(() => '')),
      numberFormats: values.map((r) => r.map(() => 'General')),
      structure: 'data_table',
      headerRowIndex: 0,
    };
  }

  const templateRows: unknown[][] = [['Unit No', 'Guest', 'Total']];
  for (let i = 0; i < 120; i += 1) templateRows.push(['', '', '']);

  function ctx(sheets: WorkbookContext['sheets']): WorkbookContext {
    return {
      activeSheetName: sheets[0]?.name ?? 'Sheet1',
      sheets,
      namedRanges: [],
      tables: [],
    };
  }

  it('names a scaffolded-but-empty sheet so a false premise is visible to the planner', () => {
    const msg = buildPlannerUserMessage(
      'some of my bank accounts got renamed, the old ones are still in the sheets',
      ctx([sheet('Jan', templateRows, 121)]),
      [],
    );
    expect(msg).toContain('Sheets with NO entered data');
    expect(msg).toContain('"Jan": 0 data rows');
  });

  it('still discloses on the promptContext path — the one real requests take', () => {
    // This branch replaces the structured sheet line entirely, so the disclosure
    // has to live outside it or it would be dropped exactly where it matters.
    const msg = buildPlannerUserMessage(
      'fix the existing rows',
      ctx([sheet('Jan', templateRows, 121)]),
      [],
      'Sheet: "Jan" | 121x3 | type: data_table',
    );
    expect(msg).toContain('Sheets with NO entered data');
  });

  it('says nothing when every sheet has real data', () => {
    const populated: unknown[][] = [
      ['Unit No', 'Guest', 'Total'],
      ['101', 'Ada', 500],
    ];
    const msg = buildPlannerUserMessage(
      'add a discount column',
      ctx([sheet('Jan', populated, 2)]),
      [],
    );
    expect(msg).not.toContain('Sheets with NO entered data');
  });
});
