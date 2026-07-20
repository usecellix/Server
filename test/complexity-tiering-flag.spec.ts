import {
  parseComplexityTieringMode,
  resolveExecutableTier,
} from '../src/excel-ai/utils/complexity-tiering-flag.util';

describe('complexity tiering feature flag', () => {
  it('defaults to off in production and full otherwise', () => {
    expect(parseComplexityTieringMode(undefined, 'production')).toBe('off');
    expect(parseComplexityTieringMode(undefined, 'development')).toBe('full');
    expect(parseComplexityTieringMode(undefined, 'test')).toBe('full');
  });

  it('parses aliases', () => {
    expect(parseComplexityTieringMode('shadow')).toBe('shadow');
    expect(parseComplexityTieringMode('tier01')).toBe('tier01');
    expect(parseComplexityTieringMode('tier0-1')).toBe('tier01');
    expect(parseComplexityTieringMode('on')).toBe('full');
    expect(parseComplexityTieringMode('false')).toBe('off');
  });

  it('shadow and off always execute as tier 3', () => {
    expect(resolveExecutableTier(0, 'off')).toBe(3);
    expect(resolveExecutableTier(2, 'shadow')).toBe(3);
    expect(resolveExecutableTier(3, 'shadow')).toBe(3);
  });

  it('tier01 only allows tiers 0 and 1', () => {
    expect(resolveExecutableTier(0, 'tier01')).toBe(0);
    expect(resolveExecutableTier(1, 'tier01')).toBe(1);
    expect(resolveExecutableTier(2, 'tier01')).toBe(3);
    expect(resolveExecutableTier(3, 'tier01')).toBe(3);
  });

  it('full preserves classified tier', () => {
    expect(resolveExecutableTier(0, 'full')).toBe(0);
    expect(resolveExecutableTier(2, 'full')).toBe(2);
    expect(resolveExecutableTier(3, 'full')).toBe(3);
  });
});
