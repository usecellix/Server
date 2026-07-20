/**
 * Thrown when a write-route turn is about to terminate with a prose answer and
 * zero SheetActions, with no clarificationsNeeded. This is a bug class (Run B),
 * not a legitimate clarification — surface as an error, never as a confident answer.
 */
export class WriteRouteNoActionError extends Error {
  readonly code = 'WRITE_ROUTE_NO_ACTION' as const;

  constructor(
    public readonly conversationId: string,
    public readonly userMessage: string,
  ) {
    super('Something went wrong applying this change — try rephrasing');
    this.name = 'WriteRouteNoActionError';
  }
}
