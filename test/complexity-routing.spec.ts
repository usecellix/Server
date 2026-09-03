import { classifyComplexity, hasCompoundSignals } from '../src/excel-ai/utils/complexity-classifier.util';

/**
 * TASKS.md #165 — how a prompt is sorted decides whether the planning pipeline
 * ever runs, so a sorting mistake is not a latency problem, it is a
 * correctness one: an under-sorted build gets a thin single-action answer with
 * no planning and no verification.
 *
 * The asymmetry that drives every choice here: over-sorting costs seconds,
 * under-sorting costs the user's actual result. So ties break upward.
 */
describe('#165 — highest tier wins, not the first pattern in the list', () => {
  it('routes "dashboard AND highlight" to lane 3, not lane 1', () => {
    // The live bug: the lane-1 `highlight` pattern sits above the lane-3
    // `dashboard` pattern, and this phrasing carries no compound signal, so
    // first-match-wins sent a dashboard build down the single-action lane.
    const result = classifyComplexity('build me a dashboard and highlight the overdue payments');
    expect(result.match?.tier).toBe(3);
  });

  it('still routes a bare highlight request to lane 1', () => {
    const result = classifyComplexity('highlight the overdue payments');
    expect(result.match?.tier).toBe(1);
  });

  it('picks the higher tier regardless of which word comes first in the sentence', () => {
    const a = classifyComplexity('highlight overdue rows and build a dashboard');
    const b = classifyComplexity('build a dashboard and highlight overdue rows');
    expect(a.match?.tier).toBe(3);
    expect(b.match?.tier).toBe(3);
  });

  it('escalates a chart request paired with a dashboard', () => {
    expect(classifyComplexity('add a chart to the dashboard').match?.tier).toBe(3);
  });

  it('leaves a lone tier-0 request alone', () => {
    // Note the pattern requires the verb before the reference ("bold B4"); the
    // reversed phrasing "make B4 bold" matches nothing and falls to the LLM
    // router. Pre-existing, unchanged here, and safe — an unmatched prompt
    // routes UP, never down.
    const result = classifyComplexity('bold B4');
    expect(result.match?.tier).toBe(0);
    expect(result.match?.actionHint).toBe('CELL_FORMAT');
  });

  it('falls through to the LLM router when nothing matches', () => {
    expect(classifyComplexity('make B4 bold').match).toBeNull();
  });

  it('keeps the action hint aligned with the winning tier', () => {
    // Tier0DirectService re-runs the pattern for capture groups, so a hint that
    // disagrees with the tier would resolve against the wrong pattern.
    const result = classifyComplexity('freeze top row');
    expect(result.match?.tier).toBe(0);
    expect(result.match?.actionHint).toBe('FREEZE_PANES');
  });
});

describe('#165 — multi-sheet build phrasings are recognised', () => {
  const buildPrompts = [
    'i need one sheet per month and a main sheet with all the details',
    'create a tab for each region',
    'make separate sheets for all departments',
    'multiple sheets for every quarter',
    'a summary sheet pulling from the others',
    'build the monthly logs plus a master sheet',
    'set up the trackers as well as an overview page',
  ];

  it.each(buildPrompts)('detects %j as compound', (prompt) => {
    expect(hasCompoundSignals(prompt)).toBe(true);
  });

  it('escalates to lane 3 when a pattern ALSO matches', () => {
    // Compound phrasing is an escalator, not an independent classifier: it
    // raises a recognised single action to lane 3.
    const result = classifyComplexity('multiple sheets for every quarter, and highlight overdue rows');
    expect(result.match?.tier).toBe(3);
    expect(result.match?.matchedBy).toBe('regex');
  });

  it('defers to the LLM router when only the compound signal fires', () => {
    // Deliberate: a null here means "this regex has no opinion", and the router
    // — which reads vague phrasing far better — decides, defaulting to 3 when
    // unsure. Hard-coding 3 was tried and reverted; see the classifier comment.
    expect(classifyComplexity('create a tab for each region').match).toBeNull();
  });

  it('does not fire on an ordinary single-sheet request', () => {
    expect(hasCompoundSignals('sort the data by date')).toBe(false);
    expect(hasCompoundSignals('make B4 bold')).toBe(false);
    expect(hasCompoundSignals('highlight expenses over 1000')).toBe(false);
  });
});

describe('#165 — the original ledger prompt still routes correctly', () => {
  const LEDGER_PROMPT =
    'i like to have multiple sheets for all months in a year, and need a main sheet it has ' +
    'all the details of the remaining sheets, in the main sheet i need to have dashboard also, ' +
    'my need to record payments and related things, which all month sheets include Unit No, ' +
    'Guest, Guest name, check in, check out, Rate per night, total amount, source, payment ' +
    'status, bank account';

  it('lands in lane 3', () => {
    expect(classifyComplexity(LEDGER_PROMPT).match?.tier).toBe(3);
  });
});
