import {
  applyBuildSpecToSubtasks,
  EXACT_HEADERS_MARKER,
  findHeaderMismatches,
  groundBuildSpec,
  parseBuildSpec,
  shouldExtractBuildSpec,
} from '../src/agents/utils/build-spec.util';
import { SpecConformanceChecker } from '../src/agents/checkers/spec-conformance.checker';
import { SpecExtractorAgent } from '../src/agents/spec-extractor.agent';
import { PlannerOutput, SubTask } from '../src/agents/types/agent.types';

// Phase 1 of LONG_PROMPT_RELIABILITY_PLAN.md — the live booking-ledger prompt.
const PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of the remaining sheets, ' +
  'in the main sheet i need to have dashboard also, my need to record payments and related things ,which all month sheets include ' +
  'Unit No, Guest, Guest name, check in, check out, Rate per night, total amount, source, payment status, bank account';

const COLUMNS = [
  'Unit No', 'Guest', 'Guest name', 'check in', 'check out',
  'Rate per night', 'total amount', 'source', 'payment status', 'bank account',
];
const MONTHS = ['January', 'February', 'March'];

const subtask = (id: string, targetSheet: string, description: string): SubTask => ({
  id, description, targetSheet, dependsOn: [], estimatedActions: 8,
});

describe('build spec (Phase 1)', () => {
  describe('groundBuildSpec', () => {
    it('keeps a spec whose columns all appear in the prompt (case/spacing-insensitive)', () => {
      const spec = groundBuildSpec(
        { sheets: [{ names: MONTHS, columns: ['Unit No', 'Check In', 'Rate Per Night'] }] },
        PROMPT,
      );
      expect(spec?.sheets).toHaveLength(1);
    });

    it('drops a sheet whose columns the user never wrote (extractor invention)', () => {
      const spec = groundBuildSpec(
        { sheets: [{ names: MONTHS, columns: ['Unit No', 'Guest', 'GSTIN', 'HSN Code'] }] },
        PROMPT,
      );
      expect(spec).toBeNull();
    });
  });

  describe('parseBuildSpec', () => {
    it('rejects garbage without throwing', () => {
      for (const bad of [null, 'x', {}, { sheets: 'no' }, { sheets: [{ names: [], columns: [] }] }]) {
        expect(parseBuildSpec(bad)).toBeNull();
      }
    });
  });

  describe('applyBuildSpecToSubtasks', () => {
    const spec = { sheets: [{ names: MONTHS, columns: COLUMNS }] };

    it('pins the header-writing month subtasks and states the columns in the description', () => {
      const out = applyBuildSpecToSubtasks(
        [
          subtask('s1', 'January', "Create sheet 'January' and write headers in row 1"),
          subtask('s2', 'Main', 'Write headers for the dashboard'),
          subtask('s3', 'January', 'Apply currency formatting'),
        ],
        spec,
      );
      expect(out[0].expectedHeaders).toEqual(COLUMNS);
      expect(out[0].description).toContain(EXACT_HEADERS_MARKER);
      expect(out[0].description).toContain('Unit No | Guest | Guest name');
      expect(out[1].expectedHeaders).toBeUndefined(); // Main is not in the spec
      expect(out[2].expectedHeaders).toBeUndefined(); // a formatting pass writes no headers
    });
  });

  describe('findHeaderMismatches', () => {
    it('allows extra computed columns but flags missing/renamed/reordered ones', () => {
      expect(findHeaderMismatches(['A', 'B', 'C'], ['A', 'Nights', 'B', 'C'])).toEqual([]);
      expect(findHeaderMismatches(['A', 'B', 'C'], ['A', 'C', 'B'])).toEqual(['C']);
      expect(findHeaderMismatches(['Unit No', 'Guest'], ['Unit', 'Guest'])).toEqual(['Unit No']);
    });
  });

  describe('shouldExtractBuildSpec', () => {
    it('skips short prompts with small plans', () => {
      expect(shouldExtractBuildSpec('sum column B', [subtask('s1', 'S', 'x')])).toBe(false);
      expect(shouldExtractBuildSpec(PROMPT, [subtask('s1', 'S', 'x')])).toBe(true);
    });
  });
});

describe('SpecConformanceChecker', () => {
  const checker = new SpecConformanceChecker();
  const pinned: SubTask = { ...subtask('s1', 'December', 'write headers'), expectedHeaders: COLUMNS };
  const headerRow = (labels: string[], sheetName = 'December') =>
    ({
      type: 'BATCH_SET',
      sheetName,
      operations: labels.map((value, i) => ({ address: `${String.fromCharCode(65 + i)}1`, value })),
    }) as never;

  it('passes the correct header row, including extra computed columns', () => {
    const withNights = [...COLUMNS.slice(0, 5), 'Nights', ...COLUMNS.slice(5)];
    expect(checker.check([{ subtask: pinned, actions: [headerRow(COLUMNS)] }]).passed).toBe(true);
    expect(checker.check([{ subtask: pinned, actions: [headerRow(withNights)] }]).passed).toBe(true);
  });

  it('fails a header row that does not match the prompt and names the missing columns', () => {
    const wrong = ['Date', 'Description', 'Amount', 'GST', 'Total'];
    const result = checker.check([{ subtask: pinned, actions: [headerRow(wrong)] }]);
    expect(result.passed).toBe(false);
    expect(result.subtaskResults[0].feedback).toContain('Unit No');
    expect(result.subtaskResults[0].issues[0].suggestion).toContain('Unit No | Guest');
  });

  it('fails the live Column1..N shape: a table created but the header row never written', () => {
    const result = checker.check([
      {
        subtask: pinned,
        actions: [
          { type: 'ADD_SHEET', sheetName: 'December' } as never,
          { type: 'CREATE_TABLE', sheetName: 'December', range: 'A1:M2', hasHeaders: true } as never,
        ],
      },
    ]);
    expect(result.passed).toBe(false);
    expect(result.subtaskResults[0].feedback).toMatch(/Column1/);
  });

  it('never judges a subtask without expectedHeaders', () => {
    const plain = subtask('s2', 'December', 'anything');
    expect(checker.check([{ subtask: plain, actions: [headerRow(['x', 'y'])] }]).passed).toBe(true);
  });

  it('reads a range write (values from A1) as well as BATCH_SET', () => {
    const rangeWrite = {
      type: 'SET_RANGE', sheetName: 'December', range: 'A1:J1', values: [COLUMNS],
    } as never;
    expect(checker.check([{ subtask: pinned, actions: [rangeWrite] }]).passed).toBe(true);
  });
});

describe('SpecExtractorAgent', () => {
  const plan = (n: number): PlannerOutput => ({
    subtasks: Array.from({ length: n }, (_, i) =>
      subtask(`s${i}`, MONTHS[i % 3], `Create sheet '${MONTHS[i % 3]}' and write headers`),
    ),
    clarificationsNeeded: [],
    confidence: 'high',
    reasoning: '',
  });
  const build = (complete: jest.Mock) =>
    new SpecExtractorAgent({ complete } as never, { openRouterModelHigh: 'm' } as never);

  it('pins headers from a grounded extraction', async () => {
    const complete = jest
      .fn()
      .mockResolvedValue(JSON.stringify({ sheets: [{ names: MONTHS, columns: COLUMNS }] }));
    const out = await build(complete).attach(PROMPT, plan(6));
    const headerSteps = out.subtasks.filter((s) => s.isDeterministicHeaderStep);
    const restSteps = out.subtasks.filter((s) => !s.isDeterministicHeaderStep);
    expect(headerSteps).toHaveLength(6);
    expect(headerSteps.every((s) => s.expectedHeaders?.length === COLUMNS.length)).toBe(true);
    expect(restSteps).toHaveLength(6);
    expect(restSteps.every((s) => s.expectedHeaders === undefined)).toBe(true);
    // Original ids preserved on the rest steps, so any dependent's dependsOn
    // reference to the original subtask id still resolves.
    expect(restSteps.map((s) => s.id).sort()).toEqual(plan(6).subtasks.map((s) => s.id).sort());
  });

  it('returns the plan untouched when the extractor throws or returns junk', async () => {
    const p = plan(6);
    expect(await build(jest.fn().mockRejectedValue(new Error('boom'))).attach(PROMPT, p)).toBe(p);
    expect(await build(jest.fn().mockResolvedValue('not json')).attach(PROMPT, p)).toBe(p);
  });

  it('makes no LLM call for a small, short request', async () => {
    const complete = jest.fn();
    await build(complete).attach('sum column B', plan(1));
    expect(complete).not.toHaveBeenCalled();
  });
});
