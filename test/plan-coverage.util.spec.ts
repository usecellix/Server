import {
  ensureReferencedSheetsPlanned,
  ensureRepeatForCoverage,
  substituteRepeatEntry,
} from '../src/agents/utils/plan-coverage.util';
import { PlanPhase, PlannerOutput, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

function subtask(id: string, targetSheet: string, description: string, dependsOn: string[] = []): SubTask {
  return { id, targetSheet, description, dependsOn, estimatedActions: 1 };
}

function context(...names: string[]): WorkbookContext {
  return {
    activeSheetName: names[0] ?? 'Sheet1',
    sheets: names.map((name) => ({
      name,
      usedRange: 'A1',
      rowCount: 0,
      columnCount: 0,
      values: [],
      formulas: [],
      numberFormats: [],
      structure: 'unknown',
      headerRowIndex: 0,
    })),
    namedRanges: [],
    tables: [],
  };
}

function plan(subtasks: SubTask[]): PlannerOutput {
  return { subtasks, clarificationsNeeded: [], confidence: 'high', reasoning: '' };
}

// The live 2026-09-10 incident's exact January subtask.
const JANUARY_DESCRIPTION =
  "Create sheet 'January' with headers: Unit No, Guest, Guest Name, Check In, Check Out, Rate Per Night, " +
  'Total Amount, Source, Payment Status, Bank Account in row 1 (A1:J1), create table tblJanuary over A1:J2, ' +
  'add dropdowns for Source (Lists!$B$3:$B$20), Payment Status (Lists!$C$3:$C$20), Bank Account ' +
  '(Lists!$D$3:$D$20) on columns H, I, J';

const monthPhase: PlanPhase = {
  id: 'p1',
  kind: 'Create the 12 monthly payment-record sheets',
  targetSheet: 'January',
  dependsOn: [],
  repeatFor: MONTHS,
};

describe('ensureRepeatForCoverage (TASKS.md #229)', () => {
  it('reproduces the live incident: 12 months planned, only January expanded → all 12 covered', () => {
    const result = ensureRepeatForCoverage(monthPhase, [subtask('s1', 'January', JANUARY_DESCRIPTION)]);

    expect(result.filled).toEqual(MONTHS.slice(1));
    expect(result.subtasks.map((s) => s.targetSheet)).toEqual(MONTHS);

    const march = result.subtasks.find((s) => s.targetSheet === 'March')!;
    expect(march.id).toBe('s1_r3');
    expect(march.description).toContain("Create sheet 'March'");
    expect(march.description).toContain('tblMarch');
    expect(march.description).not.toContain('January');
    // Dropdown sources are shared, not per-month — untouched by substitution.
    expect(march.description).toContain('Lists!$B$3:$B$20');
  });

  it('clones a multi-subtask template group and remaps its internal dependsOn', () => {
    const result = ensureRepeatForCoverage({ ...monthPhase, repeatFor: ['January', 'February'] }, [
      subtask('s1', 'January', 'Create January'),
      subtask('s2', 'January', 'Format January', ['s1']),
    ]);

    const feb = result.subtasks.filter((s) => s.targetSheet === 'February');
    expect(feb.map((s) => s.id)).toEqual(['s1_r2', 's2_r2']);
    expect(feb[1].dependsOn).toEqual(['s1_r2']);
  });

  it('only fills the entries that are actually missing', () => {
    const result = ensureRepeatForCoverage({ ...monthPhase, repeatFor: ['January', 'February', 'March'] }, [
      subtask('s1', 'January', 'Create January'),
      subtask('s2', 'February', 'Create February'),
    ]);
    expect(result.filled).toEqual(['March']);
    expect(result.subtasks).toHaveLength(3);
  });

  it('leaves a complete expansion untouched', () => {
    const subtasks = MONTHS.map((m, i) => subtask(`s${i + 1}`, m, `Create ${m}`));
    const result = ensureRepeatForCoverage(monthPhase, subtasks);
    expect(result.filled).toEqual([]);
    expect(result.subtasks).toBe(subtasks);
  });

  it('does not clone entries a single subtask already covers by name', () => {
    const result = ensureRepeatForCoverage({ ...monthPhase, repeatFor: ['January', 'February', 'March'] }, [
      subtask('s1', 'January', 'Create sheets January, February and March with headers'),
    ]);
    expect(result.filled).toEqual([]);
  });

  it('does nothing for a non-repeat phase or an empty expansion', () => {
    const plain: PlanPhase = { id: 'p2', kind: 'Main', targetSheet: 'Main', dependsOn: [] };
    expect(ensureRepeatForCoverage(plain, [subtask('s1', 'Main', 'Build Main')]).filled).toEqual([]);
    expect(ensureRepeatForCoverage(monthPhase, []).subtasks).toEqual([]);
  });
});

describe('substituteRepeatEntry', () => {
  it('rewrites the space-stripped form used in table names', () => {
    expect(substituteRepeatEntry("Create 'Jan 2026', table tblJan2026", 'Jan 2026', 'Feb 2026')).toBe(
      "Create 'Feb 2026', table tblFeb2026",
    );
  });

  it('does not rewrite a numbered entry inside a longer number', () => {
    expect(substituteRepeatEntry('Unit 1 and Unit 10', 'Unit 1', 'Unit 2')).toBe('Unit 2 and Unit 10');
  });
});

describe('ensureReferencedSheetsPlanned (TASKS.md #230)', () => {
  it('reproduces the live incident: dropdowns backed by a Lists sheet nothing creates', () => {
    const { plan: out, added } = ensureReferencedSheetsPlanned(
      plan([subtask('p1_s1', 'January', JANUARY_DESCRIPTION)]),
      context('Sheet1'),
    );

    expect(added).toEqual(['Lists']);
    const lists = out.subtasks.find((s) => s.targetSheet === 'Lists')!;
    expect(lists.description).toContain('Lists!$B$3');
    expect(lists.description).toContain('Lists!$D$3');
    expect(lists.description).toMatch(/Unassigned/);
    expect(lists.description).toMatch(/Hide the sheet/);
    expect(out.subtasks.find((s) => s.id === 'p1_s1')!.dependsOn).toEqual([lists.id]);
  });

  it('ignores sheets that already exist or that a subtask creates', () => {
    const input = plan([
      subtask('s1', 'January', 'Create January'),
      subtask('s2', 'Main', 'B6 =SUM(January!G:G), C6 =SUMIF(Data!I:I,"Paid",Data!G:G)', ['s1']),
    ]);
    const { plan: out, added } = ensureReferencedSheetsPlanned(input, context('Data'));
    expect(added).toEqual([]);
    expect(out).toBe(input);
  });

  it('treats a sheet named in prose as planned, even when it is not a targetSheet', () => {
    const { added } = ensureReferencedSheetsPlanned(
      plan([
        subtask('s1', 'January', 'Create sheets January, February and March'),
        subtask('s2', 'Main', 'B7 =SUM(February!G:G)', ['s1']),
      ]),
      context('Sheet1'),
    );
    expect(added).toEqual([]);
  });

  it('creates a formula-only referenced sheet empty rather than seeding list values', () => {
    const { plan: out, added } = ensureReferencedSheetsPlanned(
      plan([subtask('s1', 'Main', "B3 =SUM('Raw Data'!C:C)")]),
      context('Main'),
    );
    expect(added).toEqual(['Raw Data']);
    expect(out.subtasks[0].description).toMatch(/do not invent data/);
    expect(out.subtasks[0].description).not.toMatch(/Unassigned/);
  });

  it('does not mistake Excel error literals for sheet references', () => {
    const { added } = ensureReferencedSheetsPlanned(
      plan([subtask('s1', 'Main', 'Fix the #REF! and #VALUE! errors in B3')]),
      context('Main'),
    );
    expect(added).toEqual([]);
  });
});
