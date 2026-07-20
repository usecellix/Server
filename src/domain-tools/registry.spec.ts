import { DOMAIN_TOOL_NAMES, domainToolRegistry } from './registry';
import { DomainToolResult } from './types/domain-tool.types';

describe('domainToolRegistry', () => {
  it('registers every expected domain tool name', () => {
    expect(DOMAIN_TOOL_NAMES.sort()).toEqual(
      [
        'bank_recon',
        'cost_allocation',
        'gst_match',
        'ind_as_gen',
        'itc_compute',
        'tds_26as_match',
        'trial_balance_check',
      ].sort(),
    );
  });

  it('every registry entry is a DomainTool function that throws while stubbed', () => {
    for (const name of DOMAIN_TOOL_NAMES) {
      const tool = domainToolRegistry[name];
      expect(typeof tool).toBe('function');
      expect(() => tool({})).toThrow(/Not implemented/i);
    }
  });

  it('DomainToolResult requires confidence and exceptions (compile + shape contract)', () => {
    // If confidence/exceptions were optional, this assignment would still type-check
    // incorrectly — keep them required on the interface and assert at runtime here.
    const sample: DomainToolResult<{ ok: true }> = {
      data: { ok: true },
      confidence: 0.99,
      exceptions: [],
      sourceRefs: [],
    };
    expect(sample.confidence).toBeGreaterThan(0);
    expect(sample.exceptions).toEqual([]);
  });
});
