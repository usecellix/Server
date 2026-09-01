import { CellChange } from '../types/change-set.types';

/**
 * Thrown when a revert's inverse actions, dry-run against the shadow workbook, don't
 * converge back to the change set's captured `beforeState` (TASKS.md #19). Fail closed:
 * the caller must not mark the change set reverted or apply anything for real — this is
 * the A1 "no false success" rule applied to undo instead of forward execution.
 */
export class RevertVerificationError extends Error {
  readonly code = 'REVERT_VERIFICATION_FAILED' as const;

  constructor(
    public readonly changeSetId: string,
    public readonly blockingChanges: CellChange[],
  ) {
    super(
      `Revert of change set ${changeSetId} refused — ${blockingChanges.length} cell(s) would not converge to the pre-change state: ` +
        blockingChanges
          .slice(0, 5)
          .map((c) => `${c.sheet}!${c.cell} (expected ${JSON.stringify(c.before)}, would be ${JSON.stringify(c.after)})`)
          .join(', ') +
        (blockingChanges.length > 5 ? `, and ${blockingChanges.length - 5} more` : ''),
    );
    this.name = 'RevertVerificationError';
  }
}
