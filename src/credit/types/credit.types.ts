/** Billing category a request is priced against. CREDIT_SYSTEM.md CD-1. */
export type CreditActionType =
  | 'FORMULA_QA_SIMPLE'
  | 'FORMULA_QA_COMPLEX'
  | 'FORMULA_GENERATE_OR_FIX'
  | 'DATA_CLEANUP_TALLY'
  | 'GSTIN_VALIDATION_BATCH'
  | 'GST_RECONCILIATION'
  | 'ITC_COMPUTATION'
  | 'TDS_COMPLIANCE_CHECK'
  | 'AUDIT_TRAIL_PDF_EXPORT'
  | 'MIS_DASHBOARD_BUILD'
  | 'EINVOICE_VALIDATION'
  | 'BANK_RECONCILIATION_ASSIST'
  | 'TIER3_AGENTIC_BUILD';

export type CreditBucket = 'planCredits' | 'purchasedCredits' | 'oneTimeCredits';

export type PlanTier = 'free' | 'beta' | 'solo' | 'firm' | 'enterprise';

export type BillingEntityType = 'user' | 'org';

export type LedgerEntryType = 'grant' | 'purchase' | 'debit' | 'one_time_grant';

/** Resolved cost for a specific request, after per-unit quantity is applied. */
export interface ResolvedCreditCost {
  actionType: CreditActionType;
  credits: number;
}

/** Result of a gate check — whether a request may proceed to dispatch. */
export interface CreditGateResult {
  allowed: boolean;
  /** Present when allowed is false. */
  reason?: 'insufficient_balance' | 'unwired_action' | 'unknown_action_type';
  availableBalance?: number;
  requiredCredits?: number;
}

/** Result of an attempted debit. CREDIT_SYSTEM.md CD-3 / CD-6. */
export interface CreditDebitResult {
  debited: boolean;
  /** Present when debited is true — the balances after the debit. */
  balances?: {
    planCredits: number;
    purchasedCredits: number;
    oneTimeCredits: number;
  };
}
