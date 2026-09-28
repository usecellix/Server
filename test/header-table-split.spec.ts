import {
  buildHeaderTableActions,
  deterministicTableName,
  extractFullHeaderRowFromDescription,
  resolveHeaderRow,
  stripAlreadyBuiltInstructions,
  splitSpecPinnedSubtasks,
} from '../src/agents/utils/header-table-split.util';
import { findCloneGroups } from '../src/agents/utils/template-replicate.util';
import { ComputedColumnChecker } from '../src/agents/checkers/computed-column.checker';
import { SpecConformanceChecker } from '../src/agents/checkers/spec-conformance.checker';
import { SubTask } from '../src/agents/types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md — two consecutive live 12-month runs both
 * had the (single, solo) template subtask hit "max iterations (10)" or time
 * out, even with NO concurrency contention: one subtask doing create + header
 * + table + formulas + dropdowns + widths + font is too much for one Executor
 * generation. Splitting the code-certain part (create+header+table) off, built
 * deterministically with zero LLM calls, is meant to leave a lighter, more
 * reliable "rest" subtask behind.
 */

const COLUMNS = ['Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out'];

const pinned = (id: string, month: string): SubTask => ({
  id,
  description: `Create sheet '${month}' and build the payment ledger in one go: headers in A1:E1 = [Unit No, Guest, Guest Name, Check In, Check Out]; create Excel Table 'tbl${month}' over A1:E2; set formulas; add dropdowns; set widths.`,
  targetSheet: month,
  dependsOn: ['lists'],
  estimatedActions: 20,
  expectedHeaders: COLUMNS,
});

describe('buildHeaderTableActions', () => {
  it('builds ADD_SHEET + header BATCH_SET + CREATE_TABLE from expectedHeaders alone', () => {
    const actions = buildHeaderTableActions(pinned('s1', 'January')) as Array<Record<string, any>>;
    expect(actions).toHaveLength(3);

    const [sheet, headers, table] = actions;
    expect(sheet.type).toBe('ADD_SHEET');
    expect(sheet.name).toBe('January');

    expect(headers.type).toBe('BATCH_SET');
    expect(headers.sheetName).toBe('January');
    const rowOne = headers.operations.filter((op: any) => op.address.endsWith('1'));
    expect(rowOne).toEqual([
      { address: 'A1', value: 'Unit No' },
      { address: 'B1', value: 'Guest' },
      { address: 'C1', value: 'Guest Name' },
      { address: 'D1', value: 'Check In' },
      { address: 'E1', value: 'Check Out' },
    ]);
    // TASKS.md #289 — the data row is seeded too. A table spanning A1:E2 whose
    // row 2 has no cells does not report as 2 rows, and a live run refused
    // every row-2 formula write with "row 2 does not exist".
    const rowTwo = headers.operations.filter((op: any) => op.address.endsWith('2'));
    expect(rowTwo).toEqual([
      { address: 'A2', value: '' },
      { address: 'B2', value: '' },
      { address: 'C2', value: '' },
      { address: 'D2', value: '' },
      { address: 'E2', value: '' },
    ]);

    expect(table.type).toBe('CREATE_TABLE');
    expect(table.sheetName).toBe('January');
    expect(table.range).toBe('A1:E2');
    expect(table.tableName).toBe('tblJanuary');
    expect(table.hasHeaders).toBe(true);
    expect(table.showFilterButton).toBe(false);
  });

  it('handles more than 26 columns (double-letter column addressing)', () => {
    const many = Array.from({ length: 30 }, (_, i) => `Col${i + 1}`);
    const actions = buildHeaderTableActions({ ...pinned('s1', 'January'), expectedHeaders: many }) as Array<Record<string, any>>;
    const headers = actions[1];
    const rowOneOps = headers.operations.filter((op: any) => /1$/.test(op.address));
    expect(rowOneOps[25].address).toBe('Z1');
    expect(rowOneOps[26].address).toBe('AA1');
    expect(rowOneOps[29].address).toBe('AD1');
    expect(actions[2].range).toBe('A1:AD2');
  });

  it('deterministicTableName strips non-alphanumerics', () => {
    expect(deterministicTableName('January')).toBe('tblJanuary');
    expect(deterministicTableName("Guest's Log")).toBe('tblGuestsLog');
  });
});

describe('splitSpecPinnedSubtasks', () => {
  it('splits a pinned subtask in two, keeping the ORIGINAL id on the rest step', () => {
    const [header, rest] = splitSpecPinnedSubtasks([pinned('p2_s1', 'January')]);

    expect(header.isDeterministicHeaderStep).toBe(true);
    expect(header.id).toBe('hdr_January');
    expect(header.targetSheet).toBe('January');
    expect(header.expectedHeaders).toEqual(COLUMNS);
    expect(header.dependsOn).toEqual(['lists']);

    expect(rest.id).toBe('p2_s1'); // unchanged — existing dependsOn refs to it still resolve
    expect(rest.dependsOn).toEqual(['lists', 'hdr_January']);
    expect(rest.description).toContain('do NOT use ADD_SHEET, CREATE_TABLE or INSERT_COLUMN');
    expect(rest.description).toContain(pinned('p2_s1', 'January').description);
  });

  it('a subtask that other subtasks depend on keeps working — the id never changes', () => {
    const main: SubTask = {
      id: 'p3_s1', description: 'Sum months', targetSheet: 'Main',
      dependsOn: ['p2_s1', 'p2_s2'], estimatedActions: 5,
    };
    const [, jan] = splitSpecPinnedSubtasks([pinned('p2_s1', 'January')]);
    expect(jan.id).toBe('p2_s1');
    expect(main.dependsOn).toContain(jan.id); // still resolves post-split, unmodified
  });

  it('leaves an unpinned subtask (no expectedHeaders) untouched', () => {
    const plain: SubTask = { id: 's9', description: 'Format Main', targetSheet: 'Main', dependsOn: [], estimatedActions: 2 };
    expect(splitSpecPinnedSubtasks([plain])).toEqual([plain]);
  });

  it('does not re-split an already-split step (idempotent over repeated calls)', () => {
    const once = splitSpecPinnedSubtasks([pinned('p2_s1', 'January')]);
    const twice = splitSpecPinnedSubtasks(once);
    expect(twice).toHaveLength(2);
    expect(twice.filter((s) => s.isDeterministicHeaderStep)).toHaveLength(1);
  });

  it('avoids an id collision if hdr_<Sheet> already exists in the plan', () => {
    const clash: SubTask = { id: 'hdr_January', description: 'unrelated', targetSheet: 'X', dependsOn: [], estimatedActions: 1 };
    const split = splitSpecPinnedSubtasks([clash, pinned('p2_s1', 'January')]);
    const header = split.find((s) => s.isDeterministicHeaderStep);
    expect(header?.id).toBe('hdr_January_2'); // clash already holds 'hdr_January'
    expect(split.find((s) => s.id === 'hdr_January')).toBe(clash); // clash's own id, kept as-is
  });

  it('the 12-month live shape: every rest step still groups as a Phase 2 clone set', () => {
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const split = splitSpecPinnedSubtasks(months.map((m, i) => pinned(`p2_s${i + 1}`, m)));

    const headerSteps = split.filter((s) => s.isDeterministicHeaderStep);
    const restSteps = split.filter((s) => !s.isDeterministicHeaderStep);
    expect(headerSteps).toHaveLength(12);
    expect(restSteps).toHaveLength(12);

    // Header steps themselves also clone-group (belt and suspenders — even
    // though they never reach the Executor, cheap either way).
    const headerGroups = findCloneGroups(headerSteps);
    expect(headerGroups[0]?.clones).toHaveLength(11);

    // The REAL point: the "rest" steps — the ones that still cost an LLM call
    // — must ALSO still be clone-detectable, despite each now depending on a
    // DIFFERENT header-step id. This only holds because that id is built from
    // the sheet name (hdr_January), which findCloneGroups already normalizes.
    const restGroups = findCloneGroups(restSteps);
    expect(restGroups).toHaveLength(1);
    expect(restGroups[0].clones).toHaveLength(11);
  });
});

describe('ComputedColumnChecker exemption for deterministic header steps (TASKS.md #280)', () => {
  const checker = new ComputedColumnChecker();

  it('never flags a deterministic header step for a computed-sounding column with no formula', () => {
    const [header] = splitSpecPinnedSubtasks([
      { ...pinned('s1', 'January'), expectedHeaders: ['Unit No', 'Total Amount'] },
    ]);
    const actions = buildHeaderTableActions(header);
    const result = checker.check([{ subtask: header, actions }]);
    expect(result.passed).toBe(true);
  });

  it('still flags a NON-split subtask the same as before (regression guard for #270)', () => {
    const subtask: SubTask = { id: 's1', description: 'x', targetSheet: 'December', dependsOn: [], estimatedActions: 1 };
    const result = checker.check([
      {
        subtask,
        actions: [{ type: 'BATCH_SET', sheetName: 'December', operations: [{ address: 'A1', value: 'Total Amount' }] } as never],
      },
    ]);
    expect(result.passed).toBe(false);
  });
});

describe('SpecConformanceChecker sees the deterministic header step as correct by construction', () => {
  const checker = new SpecConformanceChecker();

  it('passes the header step (it wrote exactly its own expectedHeaders)', () => {
    const [header] = splitSpecPinnedSubtasks([pinned('s1', 'January')]);
    const actions = buildHeaderTableActions(header);
    expect(checker.check([{ subtask: header, actions }]).passed).toBe(true);
  });

  it('passes the rest step trivially — it writes no header row and creates no table', () => {
    const [, rest] = splitSpecPinnedSubtasks([pinned('s1', 'January')]);
    const actions = [{ type: 'SET_FORMULA', sheetName: 'January', row: 1, col: 5, formula: '=D2-C2' } as never];
    expect(checker.check([{ subtask: rest, actions }]).passed).toBe(true);
  });
});

/**
 * Live incident (TASKS.md #283): the real February subtask from a live run
 * (agent_runs run_1790097984980_dpao2kc) — its description spells out a
 * 13-COLUMN layout ("Nights" at F, "Amount Received" at K, "Balance Due" at
 * L — all computed columns interleaved with the user's 10) while
 * `expectedHeaders` (Phase 1, grounded against the user's own words only)
 * has just the 10. Building the deterministic step from `expectedHeaders`
 * alone produced a 10-column table; the "rest" step's UNCHANGED description
 * still told the Executor to write a 13-column layout, and it tried
 * INSERT_COLUMN for "Nights" into a table with no room for it — "target
 * column F already contains data" — leaving some sheets with a mix of real
 * and Excel's own placeholder ("Column1"...) headers.
 */
describe('resolveHeaderRow uses the planner\'s FULL intended layout (TASKS.md #283)', () => {
  const FEBRUARY_DESCRIPTION =
    "Create sheet 'February'. Write headers in row 1: A1='Unit No', B1='Guest', C1='Guest Name', " +
    "D1='Check In', E1='Check Out', F1='Nights', G1='Rate Per Night', H1='Total Amount', I1='Source', " +
    "J1='Payment Status', K1='Amount Received', L1='Balance Due', M1='Bank Account'. Create Excel Table " +
    "'tblFeb' over A1:M2 (header + first data row, showFilterButton: false). Set row-2 calculated-column " +
    'formulas: F2 =IF(OR(D2="",E2=""),"",E2-D2), H2 =IF(OR(F2="",G2=""),"",F2*G2), L2 =IF(H2="","",H2-N(K2)).';

  const FEBRUARY_EXPECTED = [
    'Unit No', 'Guest', 'Guest name', 'check in', 'check out',
    'Rate per night', 'total amount', 'source', 'payment status', 'bank account',
  ];

  const februarySubtask: SubTask = {
    id: 'p2_s2',
    description: FEBRUARY_DESCRIPTION,
    targetSheet: 'February',
    dependsOn: ['p1_s1'],
    estimatedActions: 26,
    expectedHeaders: FEBRUARY_EXPECTED,
  };

  it('parses the full 13-column layout, computed columns included, in their planned positions', () => {
    const full = extractFullHeaderRowFromDescription(FEBRUARY_DESCRIPTION);
    expect(full).toEqual([
      'Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out', 'Nights', 'Rate Per Night',
      'Total Amount', 'Source', 'Payment Status', 'Amount Received', 'Balance Due', 'Bank Account',
    ]);
  });

  it('resolveHeaderRow prefers the full 13-column layout over the narrower expectedHeaders', () => {
    const resolved = resolveHeaderRow(februarySubtask);
    expect(resolved).toHaveLength(13);
    expect(resolved[5]).toBe('Nights');
  });

  it('the deterministic step now builds all 13 columns — no gap left for INSERT_COLUMN to collide with', () => {
    const actions = buildHeaderTableActions(februarySubtask) as Array<Record<string, any>>;
    const [, headers, table] = actions;
    const rowOneOps = headers.operations.filter((op: any) => /[A-Z]1$/.test(op.address));
    expect(rowOneOps).toHaveLength(13);
    expect(rowOneOps[5]).toEqual({ address: 'F1', value: 'Nights' });
    expect(table.range).toBe('A1:M2');
  });

  it("the rest step's own range statement matches the FULL layout, not just expectedHeaders' 10", () => {
    const [, rest] = splitSpecPinnedSubtasks([februarySubtask]);
    expect(rest.description).toContain('A1:M2');
    expect(rest.description).toContain('Nights');
    expect(rest.description).toContain('INSERT_COLUMN'); // explicitly told not to
  });

  it('falls back to expectedHeaders alone when the description does not actually contain them (bad/partial parse)', () => {
    const mismatched: SubTask = {
      ...februarySubtask,
      description: "Create sheet 'February'. Write a KPI cell A1='Total Bookings'.",
    };
    expect(resolveHeaderRow(mismatched)).toEqual(FEBRUARY_EXPECTED);
  });

  it('extractFullHeaderRowFromDescription returns null for a sparse/non-header mention (a lone KPI cell)', () => {
    expect(extractFullHeaderRowFromDescription("Write A1='Total Bookings' and B2=5.")).toBeNull();
  });

  it('extractFullHeaderRowFromDescription returns null on a gap (not a real contiguous header row)', () => {
    expect(extractFullHeaderRowFromDescription("A1='X', C1='Y'")).toBeNull(); // B1 missing
  });
});

/**
 * Live incident (TASKS.md #284): #283 shipped a parser that handled only the
 * per-cell phrasing ("A1='Unit No', B1='Guest'"). The very NEXT live run
 * (agent_runs run_1790098924906_2b7eqv6) used a different one —
 * "Write headers in row 1 (A1:M1): Unit No, Guest, ..." — so the parse
 * returned null, the deterministic step fell back to the narrower 10-column
 * `expectedHeaders`, and the exact bug #283 set out to close reopened:
 * INSERT_COLUMN for "Nights" against an occupied column F, and January's
 * header row left as 5 real headers followed by Column10..Column2.
 */
describe('header-row parsing across every phrasing the planner actually uses (TASKS.md #284)', () => {
  const FULL_13 = [
    'Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out', 'Nights', 'Rate Per Night',
    'Total Amount', 'Source', 'Payment Status', 'Amount Received', 'Balance Due', 'Bank Account',
  ];

  it('parses the range + comma-list form (the live run that broke #283)', () => {
    const description =
      "Create sheet 'January'. Write headers in row 1 (A1:M1): Unit No, Guest, Guest Name, Check In, " +
      'Check Out, Nights, Rate Per Night, Total Amount, Source, Payment Status, Amount Received, ' +
      "Balance Due, Bank Account. Create Excel Table 'tblJanuary' over A1:M2.";
    expect(extractFullHeaderRowFromDescription(description)).toEqual(FULL_13);
  });

  it('parses the bracketed-list form', () => {
    const description =
      "Create sheet 'January' and build the payment ledger in one go: headers in A1:M1 = [Unit No, " +
      'Guest, Guest Name, Check In, Check Out, Nights, Rate Per Night, Total Amount, Source, ' +
      'Payment Status, Amount Received, Balance Due, Bank Account]; create Excel Table.';
    expect(extractFullHeaderRowFromDescription(description)).toEqual(FULL_13);
  });

  it('parses the pipe-separated form', () => {
    const description =
      "Create sheet 'December' and build the booking/payment log: headers in row 1 — Unit No | Guest | " +
      'Guest Name | Check In | Check Out | Nights | Rate Per Night | Total Amount | Source | ' +
      'Payment Status | Amount Received | Balance Due | Bank Account. Create Excel Table.';
    expect(extractFullHeaderRowFromDescription(description)).toEqual(FULL_13);
  });

  it('rejects a parse whose width contradicts the range hint (a stray comma ran the split off)', () => {
    // Range says M1 (13 columns) but only 3 entries — the split clearly went wrong.
    expect(
      extractFullHeaderRowFromDescription('Write headers in row 1 (A1:M1): Unit No, Guest, Bank Account.'),
    ).toBeNull();
  });

  it('end-to-end: the live-run subtask now builds all 13 columns instead of falling back to 10', () => {
    const liveSubtask: SubTask = {
      id: 'p2_s1',
      targetSheet: 'January',
      dependsOn: ['p1_s1'],
      estimatedActions: 26,
      expectedHeaders: [
        'Unit No', 'Guest', 'Guest name', 'check in', 'check out',
        'Rate per night', 'total amount', 'source', 'payment status', 'bank account',
      ],
      description:
        "Create sheet 'January'. Write headers in row 1 (A1:M1): Unit No, Guest, Guest Name, Check In, " +
        'Check Out, Nights, Rate Per Night, Total Amount, Source, Payment Status, Amount Received, ' +
        "Balance Due, Bank Account. Create Excel Table 'tblJanuary' over A1:M2 (header + first data row, " +
        'showFilterButton false). Set row-2 calculated-column formulas: F2 =IF(OR(D2="",E2=""),"",E2-D2).',
    };

    expect(resolveHeaderRow(liveSubtask)).toHaveLength(13);
    const actions = buildHeaderTableActions(liveSubtask) as Array<Record<string, any>>;
    const liveRowOne = actions[1].operations.filter((op: any) => /[A-Z]1$/.test(op.address));
    expect(liveRowOne).toHaveLength(13);
    expect(liveRowOne[5]).toEqual({ address: 'F1', value: 'Nights' });
    expect(actions[2].range).toBe('A1:M2');
  });
});

describe('stripAlreadyBuiltInstructions (TASKS.md #284)', () => {
  it('removes the create-sheet / write-headers / create-table sentences the header step already did', () => {
    const description =
      "Create sheet 'January'. Write headers in row 1 (A1:M1): Unit No, Guest, Bank Account. " +
      "Create Excel Table 'tblJanuary' over A1:M2 (header + first data row). " +
      'Set row-2 calculated-column formulas: F2 =IF(OR(D2="",E2=""),"",E2-D2).';
    const stripped = stripAlreadyBuiltInstructions(description);

    expect(stripped).not.toContain('Write headers');
    expect(stripped).not.toContain("Create sheet 'January'");
    expect(stripped).not.toContain('Create Excel Table');
    expect(stripped).toContain('Set row-2 calculated-column formulas');
  });

  it('NEVER strips a sentence that also carries real work (the over-stripping guard)', () => {
    const description =
      "Create sheet 'January' and set column widths B,C ~18 and apply font 'Aptos Narrow' size 10.";
    expect(stripAlreadyBuiltInstructions(description)).toBe(description);
  });

  it('leaves an unrelated description completely untouched', () => {
    const description = 'Add DATA_VALIDATION dropdowns on Source and Payment Status. Set column widths.';
    expect(stripAlreadyBuiltInstructions(description)).toBe(description);
  });

  it('falls back to the original rather than returning nothing when every sentence was already-built work', () => {
    const description = "Create sheet 'January'. Write headers in row 1: Unit No, Guest.";
    expect(stripAlreadyBuiltInstructions(description)).toBe(description);
  });

  it("the rest step's embedded description no longer contradicts the prepended notice", () => {
    const [, rest] = splitSpecPinnedSubtasks([
      {
        id: 'p2_s1',
        targetSheet: 'January',
        dependsOn: [],
        estimatedActions: 26,
        expectedHeaders: ['Unit No', 'Guest', 'Bank Account'],
        description:
          "Create sheet 'January'. Write headers in row 1 (A1:C1): Unit No, Guest, Bank Account. " +
          'Set row-2 formulas: C2 =A2.',
      },
    ]);

    expect(rest.description).toContain('do NOT use ADD_SHEET, CREATE_TABLE or INSERT_COLUMN');
    expect(rest.description).toContain('Set row-2 formulas');
    // The contradiction itself is gone, not merely warned about.
    expect(rest.description).not.toContain('Write headers in row 1');
  });
});
