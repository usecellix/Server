import { CreditActionType } from './types/credit.types';

/**
 * Single source of truth for what a billable action costs, transcribed
 * directly from cellix-pricing-v3.html's Credit Cost Map (CREDIT_SYSTEM.md
 * CD-1). Deliberately code, not a Mongo collection — CREDIT_SYSTEM_SCHEMA.md
 * §1 — so a price change requires a code review and a git history, the same
 * way the pricing doc is itself a reviewed artifact.
 *
 * Typed Record<CreditActionType, …>, so adding a billing category without
 * pricing it here fails the build — the same exhaustiveness mechanism
 * excel-ai/types/action-catalog.ts uses for SheetActionType.
 */
export type CreditCatalogEntry =
  | { kind: 'flat'; credits: number; status: 'active' | 'unwired' }
  | { kind: 'per-unit'; creditsPerUnit: number; unitSize: number; status: 'active' | 'unwired' };

export const CREDIT_COST_CATALOG: Record<CreditActionType, CreditCatalogEntry> = {
  FORMULA_QA_SIMPLE: { kind: 'flat', credits: 2, status: 'active' },
  FORMULA_QA_COMPLEX: { kind: 'flat', credits: 5, status: 'active' },
  FORMULA_GENERATE_OR_FIX: { kind: 'flat', credits: 8, status: 'active' },
  DATA_CLEANUP_TALLY: { kind: 'flat', credits: 10, status: 'active' },
  GSTIN_VALIDATION_BATCH: { kind: 'per-unit', creditsPerUnit: 3, unitSize: 100, status: 'active' },
  // Unwired per ARCHITECTURE.md AD-8 — DomainToolsModule scaffolding, gated on
  // CA sign-off. Priced now for pricing-page/future-proofing; no code path
  // can trigger these three debits yet. CREDIT_SYSTEM.md CD-5.
  GST_RECONCILIATION: { kind: 'flat', credits: 22, status: 'unwired' },
  ITC_COMPUTATION: { kind: 'flat', credits: 18, status: 'unwired' },
  TDS_COMPLIANCE_CHECK: { kind: 'flat', credits: 12, status: 'unwired' },
  AUDIT_TRAIL_PDF_EXPORT: { kind: 'flat', credits: 4, status: 'active' },
  MIS_DASHBOARD_BUILD: { kind: 'flat', credits: 28, status: 'active' },
  EINVOICE_VALIDATION: { kind: 'per-unit', creditsPerUnit: 1, unitSize: 1, status: 'active' },
  BANK_RECONCILIATION_ASSIST: { kind: 'flat', credits: 15, status: 'unwired' },
};

/** Every billing category the system knows, derived from the exhaustive catalog. */
export const ALL_CREDIT_ACTION_TYPES = Object.keys(CREDIT_COST_CATALOG) as CreditActionType[];

/** Billing categories with a live code path that can trigger their debit. */
export const WIRED_CREDIT_ACTION_TYPES = ALL_CREDIT_ACTION_TYPES.filter(
  (type) => CREDIT_COST_CATALOG[type].status === 'active',
);

/**
 * Credits required for `quantity` units of `actionType` — CD-2's per-unit
 * scaling (`ceil(quantity / unitSize) * creditsPerUnit`) for GSTIN/e-Invoice
 * validation, or the flat price for everything else. `quantity` is ignored
 * for flat entries.
 */
export function resolveCreditCost(actionType: CreditActionType, quantity = 1): number {
  const entry = CREDIT_COST_CATALOG[actionType];
  if (entry.kind === 'flat') {
    return entry.credits;
  }
  return Math.ceil(quantity / entry.unitSize) * entry.creditsPerUnit;
}

/** True when actionType has a live code path that can trigger a debit. */
export function isCreditActionWired(actionType: CreditActionType): boolean {
  return CREDIT_COST_CATALOG[actionType].status === 'active';
}
