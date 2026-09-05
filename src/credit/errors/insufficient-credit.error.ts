/**
 * Thrown by CreditLedgerService.debit when the atomic debit's floor check
 * fails — the account no longer has enough balance across all three buckets.
 * Callers on the request path should prefer CreditGateService.checkBalance
 * before dispatch (CREDIT_SYSTEM.md CD-4); this error covers the completion-
 * time debit itself, where CD-6 explains why the check and the write must be
 * one atomic operation rather than a separate read then write.
 */
export class InsufficientCreditError extends Error {
  readonly code = 'INSUFFICIENT_CREDIT' as const;

  constructor(
    public readonly billingEntityId: string,
    public readonly requiredCredits: number,
  ) {
    super(`Billing entity ${billingEntityId} has insufficient credit for a ${requiredCredits}-credit debit`);
    this.name = 'InsufficientCreditError';
  }
}
