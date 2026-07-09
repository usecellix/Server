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
});
