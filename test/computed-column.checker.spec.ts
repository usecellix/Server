import { ComputedColumnChecker } from '../src/agents/checkers/computed-column.checker';
import { Action, SubTask } from '../src/agents/types/agent.types';

/**
 * TASKS.md #270 — the live failure this exists to catch: the month sheets got
 * a "Total Amount" header and no formula anywhere, so the user typed a guest,
 * two dates and a rate and watched the cell stay blank.
 */
describe('ComputedColumnChecker', () => {
  const checker = new ComputedColumnChecker();

  const subtask = (id = 's1'): SubTask => ({
    id,
    description: "Create sheet 'December' and write headers",
    targetSheet: 'December',
    dependsOn: [],
    estimatedActions: 5,
  });

  /** The exact header row the live run wrote — all values, zero formulas. */
  const liveHeaderRow = {
    type: 'BATCH_SET',
    sheetName: 'December',
    operations: [
      { address: 'A1', value: 'Unit No' },
      { address: 'B1', value: 'Guest' },
      { address: 'C1', value: 'Guest Name' },
      { address: 'D1', value: 'Check In' },
      { address: 'E1', value: 'Check Out' },
      { address: 'F1', value: 'Rate Per Night' },
      { address: 'G1', value: 'Total Amount' },
      { address: 'H1', value: 'Source' },
      { address: 'I1', value: 'Payment Status' },
      { address: 'J1', value: 'Bank Account' },
    ],
  } as never;

  it('fails the live repro: a Total Amount header with no formula on that sheet', () => {
    const result = checker.check([{ subtask: subtask(), actions: [liveHeaderRow] }]);

    expect(result.passed).toBe(false);
    expect(result.subtaskResults[0].feedback).toContain('Total Amount');
    expect(result.subtaskResults[0].feedback).toContain('December');
    expect(result.subtaskResults[0].issues[0].suggestion).toMatch(/first DATA row/i);
  });

  it('passes once the sheet carries a formula', () => {
    const result = checker.check([
      {
        subtask: subtask(),
        actions: [
          liveHeaderRow,
          {
            type: 'BATCH_SET',
            sheetName: 'December',
            operations: [{ address: 'G2', formula: '=IF(OR(E2="",F2=""),"",E2*F2)' }],
          } as never,
        ],
      },
    ]);

    expect(result.passed).toBe(true);
  });

  it('accepts a SET_FORMULA or FILL_DOWN as the formula for that sheet', () => {
    for (const action of [
      { type: 'SET_FORMULA', sheetName: 'December', row: 1, col: 6, formula: '=E2*F2' },
      { type: 'FILL_DOWN', sheetName: 'December', sourceRange: 'G2', targetRange: 'G3:G50' },
    ]) {
      const result = checker.check([
        { subtask: subtask(), actions: [liveHeaderRow, action as never] },
      ]);
      expect(result.passed).toBe(true);
    }
  });

  it('never flags input columns — a sheet of purely typed fields is correct', () => {
    const result = checker.check([
      {
        subtask: subtask(),
        actions: [
          {
            type: 'BATCH_SET',
            sheetName: 'Lists',
            operations: [
              { address: 'A1', value: 'Source' },
              { address: 'B1', value: 'Payment Status' },
              { address: 'C1', value: 'Bank Account' },
              { address: 'D1', value: 'Guest Name' },
              { address: 'E1', value: 'Rate Per Night' },
              { address: 'F1', value: 'Amount Received' },
            ],
          } as never,
        ],
      },
    ]);

    expect(result.passed).toBe(true);
  });

  it('does not fire on a subtask that only formats or validates an existing sheet', () => {
    const result = checker.check([
      {
        subtask: subtask(),
        actions: [
          { type: 'FORMAT_RANGE', sheetName: 'December', range: 'A1:J1', format: { bold: true } } as never,
          {
            type: 'DATA_VALIDATION',
            sheetName: 'December',
            range: 'H2:H500',
            validation: { kind: 'list', listSource: 'Lists!$A$2:$A$20' },
          } as never,
        ],
      },
    ]);

    expect(result.passed).toBe(true);
  });

  it('reports each computed column once per sheet, not once per action', () => {
    const result = checker.check([
      {
        subtask: subtask(),
        actions: [
          liveHeaderRow,
          // A second pass rewriting the same header must not double-report.
          liveHeaderRow,
          {
            type: 'BATCH_SET',
            sheetName: 'January',
            operations: [
              { address: 'F1', value: 'Nights' },
              { address: 'G1', value: 'Total Amount' },
            ],
          } as never,
        ],
      },
    ]);

    expect(result.passed).toBe(false);
    // December/Total Amount once, January/Nights + January/Total Amount once each.
    expect(result.subtaskResults[0].issues).toHaveLength(3);
  });

  it('grades sheets independently — a formula on one does not excuse another', () => {
    const result = checker.check([
      {
        subtask: subtask(),
        actions: [
          liveHeaderRow,
          {
            type: 'BATCH_SET',
            sheetName: 'January',
            operations: [
              { address: 'G1', value: 'Total Amount' },
              { address: 'G2', formula: '=E2*F2' },
            ],
          } as never,
        ],
      },
    ]);

    // January is fine; December is still missing its formula.
    expect(result.passed).toBe(false);
    expect(result.subtaskResults[0].issues).toHaveLength(1);
    expect(result.subtaskResults[0].issues[0].description).toContain('December');
  });
});

/**
 * TASKS.md #298 — the blindness Phase 1.5's split introduced here.
 *
 * Before the split, one subtask wrote the headers AND the formulas, so
 * `collectWrittenHeaders` could see a "Total Amount" header and demand a
 * formula from the same subtask. After the split the header row is written by
 * the deterministic step (exempt, correctly) and the formulas belong to the
 * "rest" step — which writes no headers at all, so this checker collected
 * nothing and had no opinion on the one subtask whose job the formulas now
 * are. A live run shipped April with validations, formats and widths and zero
 * formulas, `completed: true`, and every net in the pipeline passed it.
 */
describe('ComputedColumnChecker — a split rest step (TASKS.md #298)', () => {
  const checker = new ComputedColumnChecker();
  const HEADERS = ['Unit No', 'Guest', 'Check In', 'Check Out', 'Nights', 'Total Amount'];

  const restStep = (sheet: string): SubTask => ({
    id: `p2_${sheet}`,
    targetSheet: sheet,
    dependsOn: [`hdr_${sheet}`],
    estimatedActions: 20,
    description: `Add formulas, validation and widths to '${sheet}'`,
    resolvedHeaderRow: HEADERS,
  });

  it('fails the April shape: formats and widths, no formula, no headers of its own', () => {
    const result = checker.check([
      {
        subtask: restStep('April'),
        actions: [
          { type: 'DATA_VALIDATION', sheetName: 'April', range: 'I2:I500', values: ['Paid'] },
          { type: 'FORMAT_RANGE', sheetName: 'April', range: 'A1:M1', bold: true },
          { type: 'SET_COLUMN_WIDTH', sheetName: 'April', col: 0, width: 80 },
        ] as unknown as Action[],
      },
    ]);

    expect(result.passed).toBe(false);
    expect(result.issues.some((i) => i.description.includes('Total Amount'))).toBe(true);
  });

  it('passes the same step once it writes its formula', () => {
    const result = checker.check([
      {
        subtask: restStep('March'),
        actions: [
          { type: 'SET_FORMULA', sheetName: 'March', row: 1, col: 5, formula: '=E2*D2' },
          { type: 'SET_COLUMN_WIDTH', sheetName: 'March', col: 0, width: 80 },
        ] as unknown as Action[],
      },
    ]);

    expect(result.passed).toBe(true);
  });

  it('a rest step whose pinned row is all input columns is not failed', () => {
    // The false positive that would cost a real retry on correct work.
    const result = checker.check([
      {
        subtask: {
          ...restStep('Lists'),
          resolvedHeaderRow: ['Guest', 'Bank Account', 'Rate Per Night', 'Source'],
        },
        actions: [
          { type: 'SET_COLUMN_WIDTH', sheetName: 'Lists', col: 0, width: 80 },
        ] as unknown as Action[],
      },
    ]);

    expect(result.passed).toBe(true);
  });

  it('the deterministic header step itself stays exempt (TASKS.md #280)', () => {
    const result = checker.check([
      {
        subtask: { ...restStep('July'), id: 'hdr_July', isDeterministicHeaderStep: true },
        actions: [
          { type: 'ADD_SHEET', name: 'July', sheetName: 'July' },
        ] as unknown as Action[],
      },
    ]);

    expect(result.passed).toBe(true);
  });
});
