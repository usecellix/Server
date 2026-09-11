/**
 * Thrown when a revert would produce zero inverse actions — i.e. nothing was ever
 * captured to undo. Without this, `ChangeSetService.revert()` marked the change set
 * `reverted` and returned a "successful" `{ inverseActions: [] }` even though the
 * workbook visibly didn't change back (e.g. FORMAT_MATCHING_ROWS fill color, which
 * `virtual-apply-catalog.ts` never simulates, so the shadow diff has nothing to
 * invert). Same A1 "no false success" rule as RevertVerificationError, applied to
 * the case where there was nothing to verify in the first place.
 */
export class RevertNoOpError extends Error {
  readonly code = 'REVERT_NO_OP' as const;

  constructor(
    public readonly changeSetId: string,
    public readonly irreversibleActionTypes: string[],
  ) {
    super(
      `Revert of change set ${changeSetId} refused — no inverse actions could be built` +
        (irreversibleActionTypes.length
          ? ` (${irreversibleActionTypes.join(', ')} ${irreversibleActionTypes.length === 1 ? 'is' : 'are'} not captured for revert)`
          : ''),
    );
    this.name = 'RevertNoOpError';
  }
}
