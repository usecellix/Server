import {
  describesSheetCreation,
  ensureReferencedSheetsPlanned,
  ensureRepeatForCoverage,
  ensureTargetSheetsCreated,
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

describe('describesSheetCreation (TASKS.md #262)', () => {
  it('recognizes the Planner\'s own create phrasings', () => {
    expect(describesSheetCreation("Create sheet 'January' (position after Main)", 'January')).toBe(true);
    expect(describesSheetCreation("Create the supporting sheet 'Lists' — nothing creates it", 'Lists')).toBe(true);
    expect(describesSheetCreation('Create the Main sheet with a dashboard', 'Main')).toBe(true);
    expect(describesSheetCreation('Add a new worksheet named Summary', 'Summary')).toBe(true);
  });

  it('does not treat a write-only subtask as a create', () => {
    expect(describesSheetCreation("On Main, write title in A1 ('Payments Dashboard 2026')", 'Main')).toBe(false);
    expect(describesSheetCreation('Fill Monthly Totals rows 5-10 on Main', 'Main')).toBe(false);
  });

  // The live failure: January's create names Main only as a POSITION hint.
  it('does not read one sheet\'s create as a create of another named nearby', () => {
    expect(describesSheetCreation("Create sheet 'January' (position after Main)", 'Main')).toBe(false);
  });
});

describe('ensureTargetSheetsCreated (TASKS.md #262)', () => {
  // The exact live shape: months say "Create sheet 'X'", Main subtasks only write.
  it('adds a create for a target sheet every subtask only writes to', () => {
    const { plan: out, added } = ensureTargetSheetsCreated(
      plan([
        subtask('p2_s1', 'January', "Create sheet 'January', write headers in row 1 (A1:J1)"),
        subtask('p3_s1', 'Main', "On Main, write title in A1 ('Payments Dashboard 2026')", ['p2_s1']),
        subtask('p3_s2', 'Main', 'Fill Monthly Totals rows 5-16 on Main', ['p3_s1']),
      ]),
      context('Sheet1'),
    );

    expect(added).toEqual(['Main']);
    const create = out.subtasks[0];
    expect(create.id).toBe('auto_create_1');
    expect(create.targetSheet).toBe('Main');
    // Every Main subtask must wait for it; January must not be touched.
    expect(out.subtasks.find((s) => s.id === 'p3_s1')!.dependsOn).toContain('auto_create_1');
    expect(out.subtasks.find((s) => s.id === 'p3_s2')!.dependsOn).toContain('auto_create_1');
    expect(out.subtasks.find((s) => s.id === 'p2_s1')!.dependsOn).toEqual([]);
  });

  it('adds nothing when a subtask already creates the sheet', () => {
    const { added } = ensureTargetSheetsCreated(
      plan([
        subtask('s1', 'January', JANUARY_DESCRIPTION),
        subtask('s2', 'Main', "Create the Main sheet, then write B2=SUM(January!G:G)", ['s1']),
      ]),
      context('Sheet1'),
    );
    expect(added).toEqual([]);
  });

  it('adds nothing for a sheet that already exists in the workbook', () => {
    const { added } = ensureTargetSheetsCreated(
      plan([subtask('s1', 'Main', 'On Main, write the KPI band in row 2')]),
      context('Main'),
    );
    expect(added).toEqual([]);
  });

  // ensureReferencedSheetsPlanned runs first; its auto_sheet_* creates must
  // not then be duplicated by this net.
  it('does not duplicate a create the referenced-sheet net already added', () => {
    const referenced = ensureReferencedSheetsPlanned(
      plan([subtask('s1', 'Main', 'Add dropdowns for Source (Lists!$B$3:$B$20) on column H')]),
      context('Main'),
    );
    expect(referenced.added).toEqual(['Lists']);

    const { added } = ensureTargetSheetsCreated(referenced.plan, context('Main'));
    expect(added).toEqual([]);
  });
});

describe('malformed workbook context (TASKS.md #263)', () => {
  /**
   * Live: a sheet entry whose `name` was not a string threw
   * `TypeError: name.trim is not a function` out of ensureReferencedSheetsPlanned,
   * failing the whole request AFTER the Planner had already run.
   */
  const badContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      { name: undefined },
      { name: 42 },
      { name: ['January'] },
      { name: 'Real Sheet' },
    ],
    namedRanges: [],
    tables: [],
  } as unknown as WorkbookContext;

  it('does not throw when a sheet name is not a string', () => {
    expect(() =>
      ensureReferencedSheetsPlanned(
        plan([subtask('s1', 'Main', 'Add dropdowns for Source (Lists!$B$3:$B$20)')]),
        badContext,
      ),
    ).not.toThrow();

    expect(() =>
      ensureTargetSheetsCreated(
        plan([subtask('s1', 'Main', 'On Main, write the KPI band in row 2')]),
        badContext,
      ),
    ).not.toThrow();
  });

  it('still recognizes the well-formed sheets alongside the malformed ones', () => {
    const { added } = ensureTargetSheetsCreated(
      plan([subtask('s1', 'Real Sheet', 'On Real Sheet, write a total in B2')]),
      badContext,
    );
    // 'Real Sheet' exists, so nothing to create despite its malformed siblings.
    expect(added).toEqual([]);
  });

  it('does not crash on a subtask whose targetSheet is not a string', () => {
    const malformed = plan([
      { id: 's1', targetSheet: 7, description: 'write something', dependsOn: [], estimatedActions: 1 },
    ] as unknown as SubTask[]);
    expect(() => ensureTargetSheetsCreated(malformed, context('Sheet1'))).not.toThrow();
  });
});

/**
 * The exact plan from the live 2026-09-17 failure (captured from
 * `logs/planner.log`, descriptions truncated but verbatim at the head, which
 * is what decides create-vs-write).
 *
 * The Planner produced a complete, correct-looking 18-subtask plan whose five
 * Main subtasks all WRITE to Main and none creates it. Every one of them then
 * failed in Excel with "The requested resource doesn't exist" — the user saw
 * this as "if i tap on the Accept there is no change".
 */
describe('live 2026-09-17 ledger failure — Main never created (TASKS.md #262)', () => {
  const livePlan = () =>
    plan([
      subtask('p1_s1', 'Lists', "Create sheet 'Lists' (position it at the END of the workbook, after all month and Main sheets) and populate the dropdown"),
      subtask('p2_s1', 'January', "Create sheet 'January' (position after Main), write headers in row 1: Unit No, Guest, Guest Name, Check In, Check Out, R", ['p1_s1']),
      subtask('p3_s1', 'Main', "On Main, write title in A1 ('Payments Dashboard 2026'), then KPI band in row 2: A2='Total Amount', B2=SUM(B5:B16)", ['p1_s1']),
      subtask('p3_s2', 'Main', "Fill Monthly Totals rows Jan-Jun (Main rows 5-10), one row per month, full formulas: row 5: A5='January', B5=SUM(January!G:G)", ['p1_s1', 'p2_s1']),
      subtask('p3_s3', 'Main', "Fill Monthly Totals rows Jul-Dec (Main rows 11-16), full formulas: row 11: A11='July', B11=SUM(July!G:G)", ['p1_s1']),
      subtask('p3_s4', 'Main', 'Write the CONSOLIDATED TRANSACTIONS header at Main!A18 (one blank row below the Monthly Totals table ending at row 16)', ['p3_s1', 'p3_s2', 'p3_s3']),
      subtask('p3_s5', 'Main', "Format Main: set column widths (A ~14, B-F ~14, G-K ~14), apply Indian currency numberFormat to B5:D16", ['p3_s4']),
    ]);

  it('adds the create for Main that the live plan was missing', () => {
    const { plan: out, added } = ensureTargetSheetsCreated(livePlan(), context('Sheet1'));

    expect(added).toEqual(['Main']);
    expect(out.subtasks[0].id).toBe('auto_create_1');
    expect(out.subtasks[0].targetSheet).toBe('Main');
  });

  it('makes every Main subtask wait for that create, and leaves the others alone', () => {
    const { plan: out } = ensureTargetSheetsCreated(livePlan(), context('Sheet1'));
    const byId = new Map(out.subtasks.map((s) => [s.id, s]));

    for (const id of ['p3_s1', 'p3_s2', 'p3_s3', 'p3_s4', 'p3_s5']) {
      expect(byId.get(id)!.dependsOn).toContain('auto_create_1');
    }
    // Lists and January create themselves — untouched.
    expect(byId.get('p1_s1')!.dependsOn).toEqual([]);
    expect(byId.get('p2_s1')!.dependsOn).toEqual(['p1_s1']);
  });

  it('the create lands in the first execution wave, before anything writes to Main', () => {
    const { plan: out } = ensureTargetSheetsCreated(livePlan(), context('Sheet1'));
    const create = out.subtasks.find((s) => s.id === 'auto_create_1')!;
    // No dependencies of its own — nothing can order a Main write ahead of it.
    expect(create.dependsOn).toEqual([]);
  });
});

describe('describesSheetCreation — no sheet keyword (TASKS.md #262)', () => {
  it('accepts the terse "Create <Name>" form the Planner also uses', () => {
    expect(describesSheetCreation('Create January', 'January')).toBe(true);
    expect(describesSheetCreation('Create Main', 'Main')).toBe(true);
    expect(describesSheetCreation('Create the Main dashboard', 'Main')).toBe(true);
  });

  it('still refuses a name the create verb does not govern', () => {
    // The live trap: January's create names Main only as a position hint.
    expect(describesSheetCreation("Create sheet 'January' (position after Main)", 'Main')).toBe(false);
    expect(describesSheetCreation('Write lists', 'Lists')).toBe(false);
    expect(describesSheetCreation('Add a total to Main', 'Main')).toBe(false);
  });
});
