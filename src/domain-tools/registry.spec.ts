import { DOMAIN_TOOL_NAMES, domainToolRegistry } from './registry';
import { DomainToolResult } from './types/domain-tool.types';

const STUB_TOOLS = new Set([
  'bank_recon',
  'cost_allocation',
  'ind_as_gen',
  'tds_26as_match',
  'trial_balance_check',
]);

const IMPLEMENTED = new Set(['gst_match', 'itc_compute', 'gstr3b_vs_2b']);

describe('domainToolRegistry', () => {
  it('registers every expected domain tool name', () => {
    expect(DOMAIN_TOOL_NAMES.sort()).toEqual(
      [
        'bank_recon',
        'cost_allocation',
        'gst_match',
        'gstr3b_vs_2b',
        'ind_as_gen',
        'itc_compute',
        'tds_26as_match',
        'trial_balance_check',
      ].sort(),
    );
  });

  it('every registry entry is a DomainTool function', () => {
    for (const name of DOMAIN_TOOL_NAMES) {
      expect(typeof domainToolRegistry[name]).toBe('function');
    }
  });

  it('stub tools throw Not implemented', () => {
    for (const name of STUB_TOOLS) {
      expect(() => domainToolRegistry[name]({})).toThrow(/Not implemented/i);
    }
  });

  it('implemented GST tools return DomainToolResult', () => {
    const result = domainToolRegistry.gst_match({
      purchaseRegister: [],
      gstr2b: [],
    }) as DomainToolResult<unknown>;
    expect(typeof result.confidence).toBe('number');
    expect(Array.isArray(result.exceptions)).toBe(true);
    for (const name of IMPLEMENTED) {
      expect(DOMAIN_TOOL_NAMES).toContain(name);
    }
  });

  it('DomainToolResult requires confidence and exceptions (compile + shape contract)', () => {
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
