/**
 * OpenRouter returns 400 when the model requires reasoning and the client
 * sent reasoning.effort=none (or otherwise disabled it).
 */
export function isReasoningMandatoryError(error: unknown, status?: number): boolean {
  const resolvedStatus =
    status ??
    (error && typeof error === 'object'
      ? typeof (error as { statusCode?: unknown }).statusCode === 'number'
        ? (error as { statusCode: number }).statusCode
        : typeof (error as { status?: unknown }).status === 'number'
          ? (error as { status: number }).status
          : undefined
      : undefined);

  if (resolvedStatus !== undefined && resolvedStatus !== 400) {
    return false;
  }

  const message = error instanceof Error ? error.message : String(error ?? '');
  return /reasoning is mandatory|cannot be disabled/i.test(message);
}
