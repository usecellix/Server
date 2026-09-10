import { extractJson, parseJson } from './json-extractor.util';

describe('extractJson', () => {
  it('handles bare JSON', () => {
    expect(extractJson('{"a":1}')).toBe('{"a":1}');
  });

  it('strips fenced JSON', () => {
    expect(extractJson('```json\n{"a":1}\n```')).toBe('{"a":1}');
  });

  it('strips plain fences', () => {
    expect(extractJson('```\n{"a":1}\n```')).toBe('{"a":1}');
  });

  it('extracts JSON object from prose', () => {
    expect(extractJson('Here is the plan:\n{"steps": []}')).toBe('{"steps": []}');
  });

  it('handles arrays', () => {
    expect(extractJson('[1,2,3]')).toBe('[1,2,3]');
  });

  it('repairs trailing commas', () => {
    expect(parseJson<{ a: number }>('{"a": 1,}')).toEqual({ a: 1 });
  });

  it('throws on empty input', () => {
    expect(() => extractJson('')).toThrow('empty response');
  });

  it('throws when no JSON is found', () => {
    expect(() => extractJson('no json here')).toThrow('Cannot extract JSON');
  });

  // TASKS.md #180 — a verifier response with a stray doubled leading brace
  // ("{\n{...}\n}", seen from glm-5.2) made findBalancedJsonStrings give up on
  // the whole '{'/'}' scan (the outer brace never closes), so extractJson fell
  // through to the '['/']' scan and returned the tiny, unrelated "[]" from the
  // response's own `"issues": []` field — a silent wrong match instead of a
  // thrown error. The verifier then "parsed" that into an empty array,
  // normalizeVerifierOutput saw no subtaskResults, and a correct, already
  // re-executed DELETE_CHART action got marked inconclusive and discarded.
  it('recovers the real object past a stray unmatched leading brace', () => {
    const raw =
      '{\n{"passed": true,\n"feedback": "ok",\n"issues": [],\n' +
      '"subtaskResults": [\n  {\n    "subtaskId": "s1",\n    "passed": true,\n' +
      '    "feedback": "Delete chart action targets the correct sheet.",\n    "issues": []\n  }\n]\n}';

    const parsed = parseJson<{
      passed: boolean;
      subtaskResults: Array<{ subtaskId: string; passed: boolean }>;
    }>(raw);

    expect(parsed.passed).toBe(true);
    expect(parsed.subtaskResults).toEqual([
      expect.objectContaining({ subtaskId: 's1', passed: true }),
    ]);
  });
});
