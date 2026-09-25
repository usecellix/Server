/**
 * The eight prompts LONG_PROMPT_RELIABILITY_PLAN.md §6 item 1 requires the
 * smoke harness to pass: the original booking ledger plus the seven analogous
 * ones listed in §3 (Rental, Gym, Loan/EMI, Freelance Invoice,
 * Attendance/Salary, Fleet/Fuel, Event/Wedding).
 *
 * They have been named in the plan since it was written and have NEVER been
 * run. §3 says so plainly: "Every live failure so far has come from the ONE
 * prompt we do test. Until the other 7 run, 'works for long prompts in
 * general' is an assumption, not a finding."
 *
 * Two shapes are represented deliberately:
 *   - six month-repeated prompts (the shape every bug this session came from)
 *   - one that is NOT repeated at all (Event/Wedding: four different sheets),
 *     because §2's Phase 2 risk is false-positive cloning of sheets that are
 *     genuinely different, and nothing has ever tested that end to end.
 *
 * Every column named here appears VERBATIM in its own prompt. That is not a
 * stylistic choice: Phase 1's extractor is deliberately grounded and drops any
 * column it cannot find in the prompt text, so a case whose expectations drift
 * from its prompt would be testing the harness, not the pipeline.
 * `validateCases` below enforces it offline, with no server and no model call.
 */

export interface SmokeCase {
  /** Selector for CELLIX_SMOKE_CASE. */
  id: string;
  label: string;
  prompt: string;
  /** Sheets the prompt names that must all exist when the run ends. */
  requiredSheets: string[];
  /** The dashboard/summary sheet, checked separately so its absence is legible. */
  summarySheet?: string;
  /** The column list the prompt spells out, in the order it spells it. */
  columns: string[];
  /** Columns whose cells must carry a formula, not a typed number. */
  derivedColumns: string[];
  /**
   * Sheets the column/derived checks apply to, when the prompt enumerates
   * columns for SOME of its sheets only. Defaults to every required sheet.
   */
  columnsApplyTo?: string[];
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const LEDGER_COLUMNS = [
  'Unit No', 'Guest', 'Guest name', 'check in', 'check out',
  'Rate per night', 'total amount', 'source', 'payment status', 'bank account',
];

export const SMOKE_CASES: SmokeCase[] = [
  {
    id: 'ledger',
    label: 'Booking ledger (the original)',
    prompt:
      'i like to have multiple sheets for all months in a year, and need a main sheet it has all the ' +
      'details of the remaining sheets, in the main sheet i need to have dashboard also, my need to ' +
      'record payments and related things ,which all month sheets include ' + LEDGER_COLUMNS.join(', '),
    requiredSheets: MONTHS,
    summarySheet: 'Main',
    columns: LEDGER_COLUMNS,
    derivedColumns: ['total amount'],
  },
  {
    id: 'rental',
    label: 'Property rental register',
    prompt:
      'Build me a rental register with one sheet per month for the whole year, plus a Summary sheet ' +
      'that consolidates every month and shows a dashboard. Each month sheet should have these ' +
      'columns: Property, Tenant name, Lease start, Lease end, Monthly rent, Maintenance, ' +
      'Total due, Deposit, Payment status, Notes.',
    requiredSheets: MONTHS,
    summarySheet: 'Summary',
    columns: [
      'Property', 'Tenant name', 'Lease start', 'Lease end', 'Monthly rent',
      'Maintenance', 'Total due', 'Deposit', 'Payment status', 'Notes',
    ],
    derivedColumns: ['Total due'],
  },
  {
    id: 'gym',
    label: 'Gym membership tracker',
    prompt:
      'I run a gym and want a workbook with a separate sheet for every month of the year and a ' +
      'Dashboard sheet summarising all of them. Every month sheet needs these columns: Member ID, ' +
      'Member name, Plan, Join date, Expiry date, Monthly fee, Sessions used, Amount paid, ' +
      'Balance due, Status.',
    requiredSheets: MONTHS,
    summarySheet: 'Dashboard',
    columns: [
      'Member ID', 'Member name', 'Plan', 'Join date', 'Expiry date',
      'Monthly fee', 'Sessions used', 'Amount paid', 'Balance due', 'Status',
    ],
    derivedColumns: ['Balance due'],
  },
  {
    id: 'loan',
    label: 'Loan / EMI schedule',
    prompt:
      'Create a loan tracker with one sheet per month for a full year and a Main sheet with a ' +
      'dashboard pulling from all of them. Each month sheet must contain: Loan ID, Borrower, ' +
      'Principal, Interest rate, Tenure months, EMI amount, Paid amount, Outstanding, Due date, ' +
      'Status.',
    requiredSheets: MONTHS,
    summarySheet: 'Main',
    columns: [
      'Loan ID', 'Borrower', 'Principal', 'Interest rate', 'Tenure months',
      'EMI amount', 'Paid amount', 'Outstanding', 'Due date', 'Status',
    ],
    derivedColumns: ['Outstanding'],
  },
  {
    id: 'invoice',
    label: 'Freelance invoicing',
    prompt:
      'I freelance and need an invoice book: one sheet for each month of the year plus a Summary ' +
      'sheet with a dashboard over all months. Columns on every month sheet: Invoice number, ' +
      'Client, Issue date, Due date, Hours, Rate per hour, Subtotal, Tax, Total amount, ' +
      'Payment status.',
    requiredSheets: MONTHS,
    summarySheet: 'Summary',
    columns: [
      'Invoice number', 'Client', 'Issue date', 'Due date', 'Hours',
      'Rate per hour', 'Subtotal', 'Tax', 'Total amount', 'Payment status',
    ],
    derivedColumns: ['Subtotal', 'Total amount'],
  },
  {
    id: 'attendance',
    label: 'Attendance and salary',
    prompt:
      'Set up a staff attendance and salary workbook with a sheet for every month of the year and ' +
      'a Summary sheet consolidating them into a dashboard. Each month sheet needs: Employee ID, ' +
      'Employee name, Department, Days present, Days absent, Basic pay, Overtime hours, ' +
      'Overtime pay, Deductions, Gross pay.',
    requiredSheets: MONTHS,
    summarySheet: 'Summary',
    columns: [
      'Employee ID', 'Employee name', 'Department', 'Days present', 'Days absent',
      'Basic pay', 'Overtime hours', 'Overtime pay', 'Deductions', 'Gross pay',
    ],
    derivedColumns: ['Overtime pay', 'Gross pay'],
  },
  {
    id: 'fleet',
    label: 'Fleet fuel log',
    prompt:
      'I manage a small fleet. Build a workbook with one sheet per month for the year and a ' +
      'Dashboard sheet that totals every month. Each month sheet should have these columns: ' +
      'Vehicle number, Driver, Trip date, Start odometer, End odometer, Distance, Fuel litres, ' +
      'Fuel cost, Cost per km, Notes.',
    requiredSheets: MONTHS,
    summarySheet: 'Dashboard',
    columns: [
      'Vehicle number', 'Driver', 'Trip date', 'Start odometer', 'End odometer',
      'Distance', 'Fuel litres', 'Fuel cost', 'Cost per km', 'Notes',
    ],
    derivedColumns: ['Distance', 'Cost per km'],
  },
  {
    id: 'wedding',
    label: 'Event / wedding planner (NOT a repeated shape)',
    prompt:
      'Help me plan a wedding. I want separate sheets called Guests, Budget, Vendors and Timeline, ' +
      'plus a Summary sheet with a dashboard. On the Budget sheet use these columns: Item, ' +
      'Category, Vendor, Estimated cost, Actual cost, Difference, Paid amount, Balance, ' +
      'Due date, Status.',
    requiredSheets: ['Guests', 'Budget', 'Vendors', 'Timeline'],
    summarySheet: 'Summary',
    columns: [
      'Item', 'Category', 'Vendor', 'Estimated cost', 'Actual cost',
      'Difference', 'Paid amount', 'Balance', 'Due date', 'Status',
    ],
    derivedColumns: ['Difference', 'Balance'],
    /**
     * Only the Budget sheet carries this column list; Guests/Vendors/Timeline
     * have their own shapes the prompt does not enumerate. The header and
     * derived-column checks therefore apply to Budget alone here.
     */
    columnsApplyTo: ['Budget'],
  },
];

/**
 * Offline sanity check on the case definitions themselves — no server, no
 * model call, no credits. Run with CELLIX_SMOKE_DRY_RUN=1.
 *
 * A case whose expectations do not match its own prompt would fail the run for
 * the harness's reasons rather than the pipeline's, and that is the most
 * expensive kind of false alarm: it costs a live run to discover.
 */
export function validateCases(cases: SmokeCase[] = SMOKE_CASES): string[] {
  const problems: string[] = [];
  const seen = new Set<string>();

  for (const testCase of cases) {
    const where = `[${testCase.id}]`;
    if (seen.has(testCase.id)) problems.push(`${where} duplicate case id`);
    seen.add(testCase.id);

    const prompt = testCase.prompt.toLowerCase();
    for (const column of testCase.columns) {
      // Phase 1's extractor drops any column it cannot find in the prompt.
      if (!prompt.includes(column.toLowerCase())) {
        problems.push(`${where} column "${column}" does not appear in the prompt`);
      }
    }

    for (const derived of testCase.derivedColumns) {
      if (!testCase.columns.some((c) => c.toLowerCase() === derived.toLowerCase())) {
        problems.push(`${where} derived column "${derived}" is not one of its own columns`);
      }
    }

    if (testCase.requiredSheets.length === 0) {
      problems.push(`${where} has no required sheets, so it asserts nothing about what got built`);
    }
    for (const sheet of testCase.requiredSheets) {
      if (!prompt.includes(sheet.toLowerCase()) && !prompt.includes('month')) {
        problems.push(`${where} required sheet "${sheet}" is neither named nor implied by the prompt`);
      }
    }
    if (testCase.summarySheet && testCase.requiredSheets.includes(testCase.summarySheet)) {
      problems.push(`${where} summary sheet "${testCase.summarySheet}" is also in requiredSheets`);
    }
    if (testCase.columns.length < 5) {
      problems.push(
        `${where} has only ${testCase.columns.length} columns — too narrow to be a long prompt`,
      );
    }
  }

  return problems;
}

/** The sheets a case's column/derived checks apply to. */
export function sheetsUnderColumnCheck(testCase: SmokeCase): string[] {
  return testCase.columnsApplyTo?.length ? testCase.columnsApplyTo : testCase.requiredSheets;
}
