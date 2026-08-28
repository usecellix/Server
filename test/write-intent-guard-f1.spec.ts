import { hasWriteIntent } from '../src/excel-ai/utils/write-intent-guard.util';

/**
 * F1 (2026-08-27): a pure read-only question was routed to the write planner, which
 * proposed writing KPI formulas into N2:O5 instead of answering. Cause: the
 * READ_THEN_WRITE rule matched "...and what they add up to" as "and ... add".
 */
describe('hasWriteIntent — questions must not route to the write path (F1)', () => {
  const REAL_PROMPT =
    "What's in this sheet? Give me total invoice value, total tax collected, " +
    'how many invoices are pending payment and what they add up to.';

  it('treats the reported prompt as read-only', () => {
    expect(hasWriteIntent(REAL_PROMPT)).toBe(false);
  });

  it.each([
    'What do the pending invoices add up to?',
    'How many rows are there and what do they add up to?',
    'Show me the totals and what they add up to',
    'which columns add up to the total?',
  ])('treats %j as read-only', (prompt) => {
    expect(hasWriteIntent(prompt)).toBe(false);
  });
});

describe('hasWriteIntent — compound ask-then-mutate still routes to write', () => {
  it.each([
    'Show me the pending rows and delete the empty column',
    'What is the total, then sort by amount',
    'How many are pending and highlight them in red',
    'Explain the sheet and then create a summary tab',
    'What is in column C and also add a total row',
  ])('treats %j as write intent', (prompt) => {
    expect(hasWriteIntent(prompt)).toBe(true);
  });
});

describe('hasWriteIntent — plain commands are unaffected', () => {
  it.each([
    'add a total row',
    'delete the Remarks column',
    'sort by invoice date',
    'highlight pending invoices in red',
    'Fill the Remarks column: mark every pending invoice as "Follow up"',
  ])('treats %j as write intent', (prompt) => {
    expect(hasWriteIntent(prompt)).toBe(true);
  });
});
