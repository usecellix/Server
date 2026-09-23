import {
  buildPlanSummary,
  PLAN_SUMMARY_MIN_SUBTASKS,
} from '../src/agents/utils/plan-summary.util';
import { SubTask } from '../src/agents/types/agent.types';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 5 acceptance: "summary appears for the
 * 8 known prompts; not for short prompts."
 */
const subtask = (id: string, sheet: string, deterministic = false): SubTask => ({
  id,
  description: `Build ${sheet}`,
  targetSheet: sheet,
  dependsOn: [],
  estimatedActions: 5,
  ...(deterministic ? { isDeterministicHeaderStep: true } : {}),
});

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

describe('buildPlanSummary (Phase 5, TASKS.md #288)', () => {
  it('says nothing for a short build — announcing a two-step change is just noise', () => {
    expect(buildPlanSummary([subtask('s1', 'Sheet1'), subtask('s2', 'Sheet1')])).toBeNull();
  });

  it('stays silent right up to the threshold, then speaks', () => {
    const below = Array.from({ length: PLAN_SUMMARY_MIN_SUBTASKS - 1 }, (_, i) =>
      subtask(`s${i}`, `S${i}`),
    );
    expect(buildPlanSummary(below)).toBeNull();
    expect(buildPlanSummary([...below, subtask('extra', 'Extra')])).not.toBeNull();
  });

  it('counts distinct SHEETS, not subtasks, for the headline number', () => {
    // Main gets built by several subtasks; it is still one sheet.
    const plan = [
      subtask('s1', 'Lists'),
      subtask('s2', 'Main'),
      subtask('s3', 'Main'),
      subtask('s4', 'Main'),
      subtask('s5', 'Main'),
      subtask('s6', 'Main'),
    ];
    const summary = buildPlanSummary(plan)!;
    expect(summary.sheetCount).toBe(2);
    expect(summary.subtaskCount).toBe(6);
    expect(summary.text).toContain('2 sheets');
  });

  it('the real 12-month shape reads sensibly', () => {
    const plan = [
      subtask('lists', 'Lists'),
      ...MONTHS.flatMap((m, i) => [
        subtask(`hdr_${m}`, m, true), // Phase 1.5 deterministic step
        subtask(`p2_s${i}`, m),
      ]),
      subtask('main', 'Main'),
    ];
    const summary = buildPlanSummary(plan)!;

    expect(summary.sheetCount).toBe(14); // 12 months + Lists + Main
    expect(summary.text).toContain('14 sheets');
    expect(summary.text).toMatch(/roughly \d+ minutes?/);
    expect(summary.text).toContain('stop at any point');
  });

  it('does not count the deterministic header steps toward the time estimate — they cost no model call', () => {
    const withDeterministic = [
      ...MONTHS.map((m, i) => subtask(`hdr${i}`, m, true)),
      ...MONTHS.map((m, i) => subtask(`s${i}`, m)),
    ];
    const allLlm = [...MONTHS, ...MONTHS.map((m) => `${m} B`)].map((m, i) => subtask(`s${i}`, m));

    const a = buildPlanSummary(withDeterministic)!;
    const b = buildPlanSummary(allLlm)!;
    expect(a.llmSubtaskCount).toBe(12);
    expect(b.llmSubtaskCount).toBe(24);
    expect(a.estimatedMinutes).toBeLessThan(b.estimatedMinutes);
  });

  it('never promises less than a minute', () => {
    const plan = Array.from({ length: 6 }, (_, i) => subtask(`s${i}`, `S${i}`, true));
    expect(buildPlanSummary(plan)!.estimatedMinutes).toBeGreaterThanOrEqual(1);
  });

  it('is a statement, not a question — this path must never ask for input', () => {
    const plan = MONTHS.map((m, i) => subtask(`s${i}`, m));
    expect(buildPlanSummary(plan)!.text).not.toContain('?');
  });
});
