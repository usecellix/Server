import { assessTierEscalation } from '../src/excel-ai/utils/tier-escalation.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #165 — the structural net under the word-based classifier.
 *
 * The classifier guesses a lane from the prompt's words, before any work
 * exists. This checks the guess against the work the lane actually produced,
 * which is the only reliable measure of a request's size.
 */
function cells(n: number, sheet = 'Sheet1'): SheetActionPayload[] {
  return Array.from({ length: n }, (_, i) => ({
    type: 'SET_CELL',
    sheetName: sheet,
    row: i,
    col: 0,
    value: `v${i}`,
  })) as SheetActionPayload[];
}

describe('assessTierEscalation', () => {
  it('escalates when a fast lane creates two or more sheets', () => {
    const verdict = assessTierEscalation([
      { type: 'ADD_SHEET', sheetName: 'January' },
      { type: 'ADD_SHEET', sheetName: 'February' },
    ] as SheetActionPayload[]);
    expect(verdict.escalate).toBe(true);
    expect(verdict.reason).toContain('2 sheet creations');
  });

  it('does NOT escalate a single sheet creation', () => {
    // "add a Summary sheet" is a legitimate fast-lane request.
    const verdict = assessTierEscalation([
      { type: 'ADD_SHEET', sheetName: 'Summary' },
    ] as SheetActionPayload[]);
    expect(verdict.escalate).toBe(false);
  });

  it('escalates at the staging threshold of 25 actions', () => {
    expect(assessTierEscalation(cells(24)).escalate).toBe(false);
    expect(assessTierEscalation(cells(25)).escalate).toBe(true);
  });

  it('reports why, so telemetry can separate the two causes', () => {
    expect(assessTierEscalation(cells(40)).reason).toContain('40 actions');
  });

  it('treats CREATE_SHEET and COPY_SHEET as sheet creation too', () => {
    // AD-7's drift lesson: the alias must not be a hole in the check.
    expect(
      assessTierEscalation([
        { type: 'CREATE_SHEET', sheetName: 'Q1' },
        { type: 'COPY_SHEET', sheetName: 'Q2' },
      ] as SheetActionPayload[]).escalate,
    ).toBe(true);
  });

  it('leaves an ordinary small result alone', () => {
    const verdict = assessTierEscalation([
      { type: 'CONDITIONAL_FORMAT', sheetName: 'Sheet1', range: 'B2:B50' },
    ] as SheetActionPayload[]);
    expect(verdict.escalate).toBe(false);
    expect(verdict.reason).toBeNull();
  });

  it('never escalates on empty or missing input', () => {
    expect(assessTierEscalation([]).escalate).toBe(false);
    expect(assessTierEscalation(undefined).escalate).toBe(false);
    expect(assessTierEscalation(null).escalate).toBe(false);
  });

  it('is domain-agnostic — counts shapes, never words', () => {
    // The same verdict for a hospitality build and a payroll build.
    const ledger = assessTierEscalation([
      { type: 'ADD_SHEET', sheetName: 'January' },
      { type: 'ADD_SHEET', sheetName: 'February' },
    ] as SheetActionPayload[]);
    const payroll = assessTierEscalation([
      { type: 'ADD_SHEET', sheetName: 'Engineering' },
      { type: 'ADD_SHEET', sheetName: 'Sales' },
    ] as SheetActionPayload[]);
    expect(ledger).toEqual(payroll);
  });
});
