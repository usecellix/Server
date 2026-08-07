/**
 * Phrases like "Called Cleared" / "call it Cleared" mean the VALUE is Cleared,
 * not the literal string "Called Cleared".
 */
export function extractCalledValueFromMessage(message: string): string | null {
  const text = String(message ?? '').trim();
  if (!text) return null;

  const patterns = [
    /\bcalled\s+["']([^"']+)["']/i,
    /\bcall\s+it\s+["']([^"']+)["']/i,
    /,\s*called\s+([A-Za-z0-9][\w\s/-]*?)(?:\s*[.,;]|$)/i,
    /\bcalled\s+([A-Za-z0-9][\w\s/-]*?)(?:\s*[.,;]|$)/i,
    /\bcall\s+it\s+([A-Za-z0-9][\w\s/-]*?)(?:\s*[.,;]|$)/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    const value = match?.[1]?.trim();
    if (value) return value;
  }
  return null;
}

/** Strip a leading "Called " label from a cell value when models echo the phrase. */
export function stripCalledLabel(value: unknown): unknown {
  if (typeof value !== 'string') return value;
  const match = value.trim().match(/^called\s+(.+)$/i);
  return match?.[1]?.trim() ? match[1].trim() : value;
}

export function resolveRemarkValue(
  rawValue: unknown,
  userMessage?: string,
): unknown {
  const fromMessage = userMessage ? extractCalledValueFromMessage(userMessage) : null;
  if (fromMessage) {
    const stripped = stripCalledLabel(rawValue);
    if (
      typeof stripped === 'string' &&
      stripped.trim().toLowerCase() === `called ${fromMessage}`.toLowerCase()
    ) {
      return fromMessage;
    }
    if (
      typeof rawValue === 'string' &&
      rawValue.trim().toLowerCase() === `called ${fromMessage}`.toLowerCase()
    ) {
      return fromMessage;
    }
  }
  return stripCalledLabel(rawValue);
}
