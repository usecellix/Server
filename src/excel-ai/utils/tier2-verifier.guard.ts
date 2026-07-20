/**
 * Tier 2 verification is mandatory — unlike Tier 3's shouldSkipVerifier optimization.
 * Call this guard at the Tier 2 verify call site so misuse fails loudly in dev/test.
 */
export function assertTier2VerifierMandatory(options: { usedShouldSkipVerifier: boolean }): void {
  if (options.usedShouldSkipVerifier && process.env.NODE_ENV !== 'production') {
    throw new Error('Tier 2 must never use shouldSkipVerifier — verification is mandatory here.');
  }
}
