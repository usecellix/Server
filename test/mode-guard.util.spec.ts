import { modeIsReadOnly, normalizeAssistantMode } from '../src/excel-ai/utils/mode-guard.util';

describe('normalizeAssistantMode', () => {
  it('defaults omitted mode to action (act equivalent)', () => {
    expect(normalizeAssistantMode(undefined)).toBe('action');
    expect(normalizeAssistantMode()).toBe('action');
  });

  it('accepts act as alias for action', () => {
    expect(normalizeAssistantMode('act')).toBe('action');
  });

  it('preserves ask and plan modes', () => {
    expect(normalizeAssistantMode('ask')).toBe('ask');
    expect(normalizeAssistantMode('plan')).toBe('plan');
  });

  it('treats unknown values as action', () => {
    expect(normalizeAssistantMode('unknown')).toBe('action');
  });
});

describe('modeIsReadOnly', () => {
  it('marks ask and plan as read-only', () => {
    expect(modeIsReadOnly('ask')).toBe(true);
    expect(modeIsReadOnly('plan')).toBe(true);
    expect(modeIsReadOnly('action')).toBe(false);
  });
});
