function findBalancedJsonStrings(text: string, open: '{' | '[', close: '}' | ']'): string[] {
  const results: string[] = [];
  let start = text.indexOf(open);
  while (start >= 0) {
    let depth = 0;
    let matchEnd = -1;
    for (let i = start; i < text.length; i += 1) {
      if (text[i] === open) depth += 1;
      if (text[i] === close) depth -= 1;
      if (depth === 0) {
        matchEnd = i;
        break;
      }
    }
    if (matchEnd >= 0) {
      results.push(text.slice(start, matchEnd + 1));
      start = text.indexOf(open, matchEnd + 1);
    } else {
      // This opening brace never closes before end of text — a stray/doubled
      // brace ahead of the real payload (some models emit a literal leading
      // "{\n{...}\n}"). Previously this gave up on the ENTIRE scan, which left
      // only the small unrelated "[]" from an empty `"issues": []` field as a
      // parseable candidate — extractJson happily returned that instead of
      // throwing, so the verifier silently "parsed" a malformed response into
      // an empty array and marked every subtask inconclusive, discarding a
      // correct, already-approved action (TASKS.md #180). Skip past this
      // unmatched brace and keep scanning — the real, balanced object right
      // after it still needs a chance to be found.
      start = text.indexOf(open, start + 1);
    }
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
