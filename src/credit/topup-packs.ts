/**
 * One-time credit top-up packs (cellix-pricing-v3.html's "credit top-up
 * packs available" feature). Deliberately NOT part of `CreditActionType`/
 * `CREDIT_COST_CATALOG` — a top-up is a purchase (adds to `purchasedCredits`
 * via `CreditLedgerService.addPurchasedCredits`), not a billable action that
 * debits a balance.
 *
 * Sizes re-derived alongside the credit-system-v2 repricing session: the
 * original packs (100/350/800 credits at ₹149/₹399/₹799) were sized against
 * Solo's old 1,100-credit/month allowance. Scaled ~3x here to stay
 * proportional now that Solo grants 3,000/month — a top-up should still
 * meaningfully extend a month, not just add ~10% to a much larger balance.
 */
export const TOPUP_PACKS = {
  small: { credits: 300, priceInr: 149 },
  medium: { credits: 1000, priceInr: 399 },
  large: { credits: 2200, priceInr: 799 },
} as const;

export type TopupPackId = keyof typeof TOPUP_PACKS;
