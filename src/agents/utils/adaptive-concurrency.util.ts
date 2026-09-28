/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 3 — how many Executor calls a wave may
 * have in flight at once, adjusted from what the provider actually tolerates.
 *
 * TASKS.md #275 capped this at a flat 3 after unbounded fan-out (12 month
 * sheets at once) tripped the provider's concurrency ceiling. A flat cap is a
 * guess in both directions: too high on a bad day, and needlessly slow on a
 * good one. This is the standard additive-increase / multiplicative-decrease
 * rule — halve the moment the provider pushes back, and earn width back only
 * after a run of clean successes.
 *
 * Deliberately conservative on the way up (one extra slot per
 * `increaseAfterSuccesses`) and sharp on the way down (halve immediately), so
 * a provider under strain is relieved fast and probed back slowly.
 */
export interface AdaptiveConcurrencyOptions {
  start?: number;
  floor?: number;
  ceiling?: number;
  /** Consecutive successes needed before widening by one. */
  increaseAfterSuccesses?: number;
}

export class AdaptiveConcurrency {
  private limitValue: number;
  private consecutiveSuccesses = 0;
  private readonly floor: number;
  private readonly ceiling: number;
  private readonly increaseAfterSuccesses: number;

  constructor(options: AdaptiveConcurrencyOptions = {}) {
    this.floor = Math.max(1, options.floor ?? 1);
    this.ceiling = Math.max(this.floor, options.ceiling ?? 6);
    this.increaseAfterSuccesses = Math.max(1, options.increaseAfterSuccesses ?? 4);
    this.limitValue = Math.min(this.ceiling, Math.max(this.floor, options.start ?? 3));
  }

  get limit(): number {
    return this.limitValue;
  }

  /** A call came back clean. Widens by one only after a run of them. */
  recordSuccess(): void {
    this.consecutiveSuccesses += 1;
    if (this.consecutiveSuccesses >= this.increaseAfterSuccesses && this.limitValue < this.ceiling) {
      this.limitValue += 1;
      this.consecutiveSuccesses = 0;
    }
  }

  /**
   * The provider pushed back (a transient fault). Halves immediately and
   * resets the success run — earning width back has to start over.
   */
  recordTransientFailure(): void {
    this.consecutiveSuccesses = 0;
    this.limitValue = Math.max(this.floor, Math.floor(this.limitValue / 2));
  }
}
