import {
  extractFullHeaderRowFromDescription,
  findDelimitedListContaining,
  resolveHeaderRow,
} from '../src/agents/utils/header-table-split.util';
import { repairListValidationSources } from '../src/agents/utils/list-validation-repair.util';
import { Action, SubTask } from '../src/agents/types/agent.types';

/**
 * Live run_1790332994308_etr7oxl, built from its real data.
 *
 * #332: every month sheet built 10 columns wide under a 13-column formula
 *       layout — Total Amount landed in "source", Balance Due outside the table.
 * #334: dropdowns pointed at the wrong Lists columns — bank accounts under
 *       Source, "Paid" missing from Payment Status, an empty list for Bank.
 */

const USER = [
  'Unit No', 'Guest', 'Guest name', 'check in', 'check out',
  'Rate per night', 'total amount', 'source', 'payment status', 'bank account',
];
const FULL = [
  'Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out', 'Nights', 'Rate Per Night',
  'Total Amount', 'Source', 'Payment Status', 'Amount Received', 'Balance Due', 'Bank Account',
];

// Verbatim from the run's p2_s12 description.
const LIVE_DESCRIPTION =
  "Create sheet 'December' and build the payment-tracking template: (1) write headers in row 1, columns A-M: " +
  'Unit No | Guest | Guest Name | Check In | Check Out | Nights | Rate Per Night | Total Amount | Source | ' +
  "Payment Status | Amount Received | Balance Due | Bank Account; (2) create an Excel Table 'tblDecember' over " +
  'A1:M2 (header + first data row, showFilterButton: false); (3) set calculated-column formulas in row 2: ' +
  'F2 =IF(OR(D2="",E2=""),"",E2-D2) for Nights, H2 =IF(OR(F2="",G2=""),"",F2*G2) for Total Amount.';

const month = (description: string): SubTask => ({
  id: 'p2_s12', targetSheet: 'December', description, dependsOn: [], estimatedActions: 20, expectedHeaders: USER,
});

describe('#332 — the header row is the planner\'s full layout, however it is phrased', () => {
  it('reads the live description as all 13 columns (the list ends at its ";")', () => {
    expect(extractFullHeaderRowFromDescription(LIVE_DESCRIPTION)).toEqual(FULL);
    expect(resolveHeaderRow(month(LIVE_DESCRIPTION))).toEqual(FULL);
  });

  it('finds the layout from a pipe list even when no lead-in phrase is recognised', () => {
    const novel =
      'Lay out the template like this => ' + FULL.join(' | ') + '. Then add formulas for Nights and Total Amount.';
    expect(resolveHeaderRow(month(novel))).toEqual(FULL);
  });

  it('never takes a list that is missing one of the user\'s columns', () => {
    const missingSource = FULL.filter((c) => c !== 'Source').join(' | ');
    expect(findDelimitedListContaining(`layout: ${missingSource}.`, USER)).toBeNull();
    expect(resolveHeaderRow(month(`layout: ${missingSource}.`))).toEqual(USER);
  });
});

describe('#334 — dropdowns point at the list their column is named for', () => {
  // The run's real Lists layout: row-1 headers, values from row 2.
  const lists = {
    type: 'BATCH_SET',
    sheetName: 'Lists',
    operations: [
      { address: 'A1', value: 'Source' },
      ...['Direct', 'Airbnb', 'Booking.com', 'MakeMyTrip', 'Goibibo', 'Agoda', 'OYO', 'Other'].map((v, i) => ({
        address: `A${i + 2}`, value: v,
      })),
      { address: 'B1', value: 'Payment Status' },
      { address: 'B2', value: 'Paid' }, { address: 'B3', value: 'Partial' }, { address: 'B4', value: 'Unpaid' },
      { address: 'C1', value: 'Bank Account' },
      { address: 'C2', value: 'Account 1' }, { address: 'C3', value: 'Account 2' },
    ],
  } as unknown as Action;
  const header = {
    type: 'BATCH_SET',
    sheetName: 'December',
    operations: FULL.map((value, i) => ({ address: `${String.fromCharCode(65 + i)}1`, value })),
  } as unknown as Action;
  const dropdown = (range: string, listSource: string, promptTitle: string) =>
    ({ type: 'DATA_VALIDATION', sheetName: 'December', range, validation: { kind: 'list', listSource, promptTitle } }) as unknown as Action;
  const src = (a: Action) => (a as unknown as { validation: { listSource: string } }).validation.listSource;

  it('re-points the three live dropdowns at Payment Status, Source and Bank Account', () => {
    const status = dropdown('J2:J500', 'Lists!$B$3:$B$10', 'Payment Status');
    const source = dropdown('I2:I500', 'Lists!$C$3:$C$12', 'Source');
    const bank = dropdown('M2:M500', 'Lists!$D$3:$D$6', 'Bank Account');
    const wave = [status, source, bank];

    const repairs = repairListValidationSources(wave, [lists, header, ...wave]);

    expect(src(status)).toBe('Lists!$B$2:$B$4'); // "Paid" included again
    expect(src(source)).toBe('Lists!$A$2:$A$9'); // booking sources, not bank accounts
    expect(src(bank)).toBe('Lists!$C$2:$C$3'); // a real list, not an empty column
    expect(repairs).toHaveLength(3);
  });

  it('leaves a correct dropdown, and one named for no list, exactly as they were', () => {
    const right = dropdown('J2:J500', 'Lists!$B$2:$B$4', 'Payment Status');
    const unknown = dropdown('A2:A500', 'Lists!$A$2:$A$9', 'Something Else');
    const context = [lists, { ...header, operations: [{ address: 'A1', value: 'Unit No' }] } as unknown as Action];
    expect(repairListValidationSources([right, unknown], [...context, header])).toEqual([]);
    expect(src(unknown)).toBe('Lists!$A$2:$A$9');
  });
});

/**
 * TASKS.md #336 — live run_1790338103271_z73lf81. The step never NAMED its
 * extra columns — its only list is the user's 10 — but it used them: formulas
 * in F2/H2/L2, dropdowns in I/J/M, widths to M. The sheet was built without
 * Nights; the user: "rate per night is there but total days is not".
 */
describe('#336 — extra columns the step uses but never names', () => {
  const LIVE =
    'Row 2 formulas: F2 =IF(OR(D2="",E2=""),"",E2-D2); H2 =IF(OR(F2="",G2=""),"",F2*G2); L2 =IF(H2="","",H2-N(K2)). ' +
    'Data validation: J2:J500 list from Lists!$B$3:$B$20 (Payment Status), I2:I500 list from Lists!$A$3:$A$20 (Source), ' +
    'M2:M500 list from Lists!$C$3:$C$20 (Bank Account). Column widths: A 10, B 12, C 20, D 12, E 12, F 8, G 14, H 14, ' +
    'I 14, J 14, K 16, L 14, M 18. Font Aptos Narrow 10 over A1:M2. EXACT COLUMN HEADERS (verbatim, in this order — you ' +
    'may add computed columns after or between them, but never rename, drop or reorder these): Unit No | Guest | ' +
    'Guest name | check in | check out | Rate per night | total amount | source | payment status | bank account';

  it('builds Nights after check out, and Amount Received + Balance Due after payment status', () => {
    expect(resolveHeaderRow(month(LIVE))).toEqual([
      'Unit No', 'Guest', 'Guest name', 'check in', 'check out', 'Nights', 'Rate per night',
      'total amount', 'source', 'payment status', 'Amount Received', 'Balance Due', 'bank account',
    ]);
  });

  it('never widens a step that only uses the user\'s own columns', () => {
    const narrow = 'Row 2 formula: G2 =F2*2. Column widths: A 10, J 18. Font over A1:J2.';
    expect(resolveHeaderRow(month(narrow))).toEqual(USER);
  });

  it('never widens when the references do not match the ledger layout exactly', () => {
    const eleven = 'Row 2 formula: K2 =G2*2. Font over A1:K2.';
    expect(resolveHeaderRow(month(eleven))).toEqual(USER);
  });

  it('ignores other sheets\' references when measuring this sheet', () => {
    const listsOnly = 'Dropdown in J2:J500 from Lists!$M$2:$M$20. Font over A1:J2.';
    expect(resolveHeaderRow(month(listsOnly))).toEqual(USER);
  });
});
