import { PLANNER_RULES_BODY, PLANNER_SYSTEM_PROMPT } from '../src/agents/prompts/planner.prompt';

/**
 * TASKS.md #191 split `PLANNER_SYSTEM_PROMPT`'s 130-line rule body into its
 * own export (`PLANNER_RULES_BODY`) so two-pass planning's phase-expansion
 * prompt can reuse it, rather than duplicating tuned, incident-driven text or
 * dropping it for the very requests two-pass planning exists to handle. This
 * pins that the split is lossless — `PLANNER_SYSTEM_PROMPT` must remain
 * exactly what it was before the split, not merely "close enough".
 */
describe('planner.prompt.ts — rule-body split is lossless', () => {
  it('PLANNER_SYSTEM_PROMPT ends with PLANNER_RULES_BODY verbatim', () => {
    expect(PLANNER_SYSTEM_PROMPT.endsWith(PLANNER_RULES_BODY)).toBe(true);
  });

  it('PLANNER_RULES_BODY starts with "Rules:" and contains the known critical rules', () => {
    expect(PLANNER_RULES_BODY.startsWith('Rules:')).toBe(true);
    // A handful of the most load-bearing, incident-driven rules — if any of
    // these vanished during the split, this is the test that should catch it.
    expect(PLANNER_RULES_BODY).toContain('YEARLY MONTHLY LEDGER');
    expect(PLANNER_RULES_BODY).toContain('estimatedActions IS A TOKEN BUDGET');
    expect(PLANNER_RULES_BODY).toContain('NUMBER / DATE FORMAT PRESERVATION');
    expect(PLANNER_RULES_BODY).toContain('MULTI-CLAUSE REQUESTS');
  });

  it('PLANNER_SYSTEM_PROMPT still contains its own output-schema preamble before the rules', () => {
    const rulesIndex = PLANNER_SYSTEM_PROMPT.indexOf('Rules:');
    const schemaIndex = PLANNER_SYSTEM_PROMPT.indexOf('"subtasks"');
    expect(schemaIndex).toBeGreaterThan(-1);
    expect(rulesIndex).toBeGreaterThan(schemaIndex);
  });
});
