function findJsonObjectStrings(text: string): string[] {
  const results: string[] = [];
  let start = text.indexOf('{');
  while (start >= 0) {
    let depth = 0;
    let closed = false;
    for (let i = start; i < text.length; i += 1) {
      if (text[i] === '{') depth += 1;
      if (text[i] === '}') depth -= 1;
      if (depth === 0) {
        results.push(text.slice(start, i + 1));
        start = text.indexOf('{', i + 1);
        closed = true;
        break;
      }
    }
    if (!closed) break;
  }
  return results.reverse();
}

function findJsonArrayStrings(text: string): string[] {
  const results: string[] = [];
  let start = text.indexOf('[');
  while (start >= 0) {
    let depth = 0;
    let closed = false;
    for (let i = start; i < text.length; i += 1) {
      if (text[i] === '[') depth += 1;
      if (text[i] === ']') depth -= 1;
      if (depth === 0) {
        results.push(text.slice(start, i + 1));
        start = text.indexOf('[', i + 1);
        closed = true;
        break;
      }
    }
    if (!closed) break;
  }
  return results.reverse();
}

function repairJsonString(raw: string): string {
  let s = raw.trim();
  s = s.replace(/^\uFEFF/, '');
  s = s.replace(/[“”]/g, '"').replace(/[‘’]/g, "'");
  s = s.replace(/,\s*([}\]])/g, '$1');
  s = s.replace(/\/\*[\s\S]*?\*\/|\/\/.*$/gm, '');
  return s;
}

function tryParseJson(raw: string): unknown | null {
  const attempts = [raw, repairJsonString(raw)];
  for (const candidate of attempts) {
    try {
      const parsed = JSON.parse(candidate);
      if (parsed && typeof parsed === 'object') {
        return parsed;
      }
    } catch {
      // continue
    }
  }
  return null;
}

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
  const trimmed = raw.trim();
  if (!trimmed) {
    throw new Error('Empty LLM response');
  }

  const candidates: string[] = [];

  const fencePattern = /```(?:json)?\s*([\s\S]*?)```/gi;
  let fenceMatch: RegExpExecArray | null;
  while ((fenceMatch = fencePattern.exec(trimmed)) !== null) {
    candidates.push(fenceMatch[1].trim());
  }

  candidates.push(trimmed);
  candidates.push(...findJsonObjectStrings(trimmed));

  for (const candidate of candidates) {
    const parsed = tryParseJson(candidate);
    if (parsed !== null) {
      return parsed as T;
    }
  }

  throw new Error('No valid JSON object found in LLM response');
}

/** Parse executor responses that may be a wrapper object, a bare actions array, or a single action. */
export function parseExecutorPayload(raw: string): Record<string, unknown> | null {
  const trimmed = raw.trim();
  if (!trimmed) return null;

  const candidates: string[] = [];

  const fencePattern = /```(?:json)?\s*([\s\S]*?)```/gi;
  let fenceMatch: RegExpExecArray | null;
  while ((fenceMatch = fencePattern.exec(trimmed)) !== null) {
    candidates.push(fenceMatch[1].trim());
  }

  candidates.push(trimmed);
  candidates.push(...findJsonObjectStrings(trimmed));
  candidates.push(...findJsonArrayStrings(trimmed));

  for (const candidate of candidates) {
    const parsed = tryParseJson(candidate);
    const wrapped = wrapExecutorPayload(parsed);
    if (wrapped) return wrapped;
  }

  return null;
}
