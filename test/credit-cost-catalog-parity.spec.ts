import {
  ALL_CREDIT_ACTION_TYPES,
  CREDIT_COST_CATALOG,
  WIRED_CREDIT_ACTION_TYPES,
  isCreditActionWired,
  resolveCreditCost,
} from '../src/credit/credit-cost-catalog';

/**
 * CREDIT_SYSTEM.md CD-1 / CD-5. The catalog is typed
 * Record<CreditActionType, …>, so tsc already enforces that every billing
 * category is priced. This suite pins the catalog's values to
 * cellix-pricing-v3.html's Credit Cost Map and asserts the CD-5 "unwired"
 * entries stay explicitly marked rather than silently reachable or silently
 * dropped — the same pattern action-catalog-parity.spec.ts uses for
 * SheetActionType exhaustiveness.
 */
describe('credit cost catalog parity', () => {
  it('prices every billing category from the pricing doc', () => {
    const expected: Record<string, number> = {
      FORMULA_QA_SIMPLE: 2,
      FORMULA_QA_COMPLEX: 5,
      FORMULA_GENERATE_OR_FIX: 8,
      DATA_CLEANUP_TALLY: 10,
      AUDIT_TRAIL_PDF_EXPORT: 4,
      MIS_DASHBOARD_BUILD: 28,
      GST_RECONCILIATION: 22,
      ITC_COMPUTATION: 18,
      TDS_COMPLIANCE_CHECK: 12,
      BANK_RECONCILIATION_ASSIST: 15,
    };
    for (const [actionType, credits] of Object.entries(expected)) {
      const entry = CREDIT_COST_CATALOG[actionType as keyof typeof CREDIT_COST_CATALOG];
      expect(entry.kind).toBe('flat');
      if (entry.kind === 'flat') {
        expect(entry.credits).toBe(credits);
      }
    }
  });

  it('prices the two per-unit categories per CD-2', () => {
    expect(CREDIT_COST_CATALOG.GSTIN_VALIDATION_BATCH).toEqual({
      kind: 'per-unit',
      creditsPerUnit: 3,
      unitSize: 100,
      status: 'active',
    });
    expect(CREDIT_COST_CATALOG.EINVOICE_VALIDATION).toEqual({
      kind: 'per-unit',
      creditsPerUnit: 1,
      unitSize: 1,
      status: 'active',
    });
  });

  it('marks exactly the four AD-8-gated domain-tool categories as unwired (CD-5)', () => {
    const unwired = ALL_CREDIT_ACTION_TYPES.filter((type) => !isCreditActionWired(type));
    expect(unwired.sort()).toEqual(
      ['GST_RECONCILIATION', 'ITC_COMPUTATION', 'TDS_COMPLIANCE_CHECK', 'BANK_RECONCILIATION_ASSIST'].sort(),
    );
    expect(unwired.length + WIRED_CREDIT_ACTION_TYPES.length).toBe(ALL_CREDIT_ACTION_TYPES.length);
  });

  it('scales per-unit cost with ceil(quantity / unitSize) — CD-2', () => {
    expect(resolveCreditCost('GSTIN_VALIDATION_BATCH', 340)).toBe(12); // ceil(340/100) * 3
    expect(resolveCreditCost('GSTIN_VALIDATION_BATCH', 100)).toBe(3);
    expect(resolveCreditCost('GSTIN_VALIDATION_BATCH', 101)).toBe(6);
    expect(resolveCreditCost('EINVOICE_VALIDATION', 1)).toBe(1);
    expect(resolveCreditCost('EINVOICE_VALIDATION', 5)).toBe(5);
  });

  it('ignores quantity for flat-priced categories', () => {
    expect(resolveCreditCost('FORMULA_GENERATE_OR_FIX', 99)).toBe(8);
  });
});
