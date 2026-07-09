import { parseJson } from './json-extractor.util';

function wrapExecutorPayload(parsed: unknown): Record<string, unknown> | null {
  if (!parsed || typeof parsed !== 'object') return null;

  if (Array.isArray(parsed)) {
    return { actions: parsed, isDone: true };
  }

  const record = parsed as Record<string, unknown>;
  if (Array.isArray(record.actions)) {
    return record;
  }

  if (record.type && typeof record.type === 'string') {
    return { actions: [record], isDone: true };
  }

  return record;
}

export function parseAgentJson<T>(raw: string): T {
  return parseJson<T>(raw);
}

/** Parse executor responses that may be a wrapper object, a bare actions array, or a single action. */
export function parseExecutorPayload(raw: string): Record<string, unknown> | null {
  try {
    const parsed = parseJson(raw);
    return wrapExecutorPayload(parsed);
  } catch {
    return null;
  }
}
