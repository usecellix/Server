function findBalancedJsonStrings(text: string, open: '{' | '[', close: '}' | ']'): string[] {
  const results: string[] = [];
  let start = text.indexOf(open);
  while (start >= 0) {
    let depth = 0;
    let closed = false;
    for (let i = start; i < text.length; i += 1) {
      if (text[i] === open) depth += 1;
      if (text[i] === close) depth -= 1;
      if (depth === 0) {
        results.push(text.slice(start, i + 1));
        start = text.indexOf(open, i + 1);
        closed = true;
        break;
      }
    }
    if (!closed) break;
  }
  return results.reverse();
}

function repairJsonString(raw: string): string {
  let value = raw.trim();
  value = value.replace(/^\uFEFF/, '');
  value = value.replace(/[“”]/g, '"').replace(/[‘’]/g, "'");
  value = value.replace(/,\s*([}\]])/g, '$1');
  value = value.replace(/\/\*[\s\S]*?\*\/|\/\/.*$/gm, '');
  return value;
}

function parseCandidate(raw: string): { source: string; parsed: unknown } | null {
  const attempts = [raw, repairJsonString(raw)];
  for (const attempt of attempts) {
    try {
      const parsed = JSON.parse(attempt);
      if (parsed && typeof parsed === 'object') {
        return { source: attempt, parsed };
      }
    } catch {
      // continue
    }
  }
  return null;
}

/**
 * Robustly extracts a parseable JSON string from raw LLM output.
 */
export function extractJson(raw: string): string {
  const trimmed = raw?.trim() ?? '';
  if (!trimmed) {
    throw new Error('LLM returned an empty response');
  }

  const candidates: string[] = [];
  const fencePattern = /```(?:json)?\s*([\s\S]*?)```/gi;
  let fenceMatch: RegExpExecArray | null;
  while ((fenceMatch = fencePattern.exec(trimmed)) !== null) {
    candidates.push(fenceMatch[1].trim());
  }

  candidates.push(trimmed);
  candidates.push(...findBalancedJsonStrings(trimmed, '{', '}'));
  candidates.push(...findBalancedJsonStrings(trimmed, '[', ']'));

  for (const candidate of candidates) {
    const parsed = parseCandidate(candidate);
    if (parsed) {
      return parsed.source;
    }
  }

  throw new Error(`Cannot extract JSON from LLM response. First 300 chars: ${trimmed.slice(0, 300)}`);
}

/**
 * Extracts and parses JSON from raw LLM output in one step.
 */
export function parseJson<T = unknown>(raw: string): T {
  const jsonStr = extractJson(raw);
  try {
    return JSON.parse(jsonStr) as T;
  } catch (error: unknown) {
    throw new Error(
      `JSON.parse failed after extraction.\nExtracted: ${jsonStr.slice(0, 300)}\nOriginal error: ${String(error)}`,
    );
  }
}
