/**
 * Thrown when a checkpoint restore cannot be completed safely (TASKS.md #29).
 * Fail closed: nothing is written to the database and the caller must not
 * report success — same A1 "no false success" principle RevertVerificationError
 * applies to a single revert, extended to a chain of them.
 */
export class RestoreVerificationError extends Error {
  readonly code = 'RESTORE_VERIFICATION_FAILED' as const;

  constructor(
    public readonly checkpointId: string,
    public readonly blockingChangeSetId: string,
    public readonly reason: string,
  ) {
    super(
      `Restore of checkpoint ${checkpointId} refused — change set ${blockingChangeSetId} ` +
        `cannot be safely reverted: ${reason}`,
    );
    this.name = 'RestoreVerificationError';
  }
}
