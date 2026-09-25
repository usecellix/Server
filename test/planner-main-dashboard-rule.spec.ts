import { PLANNER_SYSTEM_PROMPT } from '../src/agents/prompts/planner.prompt';

/**
 * TASKS.md #324 / #333 — the same live user, twice:
 *  - a Main listing every booking: "shows all months — it needs to be single
 *    month in that and total to be there" (#324: one row per month + Total);
 *  - a Main with only those totals: the prompt asks for "all the details of
 *    the remaining sheets" and "below it is blank" (#333: the details belong
 *    on the dashboard too, below the summary).
 * So Main is a per-month dashboard that ALSO carries every booking.
 */
describe('planner prompt — Main is a dashboard that also carries all the details', () => {
  it('lays Main out as one row per month closed by a TOTAL row', () => {
    expect(PLANNER_SYSTEM_PROMPT).toContain('Main IS A DASHBOARD that ALSO carries all the details');
    expect(PLANNER_SYSTEM_PROMPT).toContain("A17 'Total', B17 =SUM(B5:B16)");
  });

  it('"all the details" wording asks for the consolidated section below the summary', () => {
    expect(PLANNER_SYSTEM_PROMPT).toMatch(
      /CONSOLIDATED TRANSACTIONS HEADER \(required whenever the user's wording says Main should have "all details"/,
    );
    expect(PLANNER_SYSTEM_PROMPT).not.toContain('does NOT qualify');
  });

  it('keeps the TOTAL row out of the chart and puts the list below it', () => {
    expect(PLANNER_SYSTEM_PROMPT).toContain('NEVER include the TOTAL row');
    expect(PLANNER_SYSTEM_PROMPT).toContain('ALWAYS row 19 (Main!A19)');
  });
});
