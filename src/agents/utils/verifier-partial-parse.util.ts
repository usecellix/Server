import { VerifierIssue, VerifierSubtaskResult } from '../types/agent.types';

/**
 * Extract complete subtaskResults objects from a truncated / malformed verifier
 * response. Entries that appear fully closed in the JSON text are preserved;
 * the truncated tail is ignored (caller marks missing IDs as inconclusive).
 */
export function salvageVerifierSubtaskResults(raw: string): VerifierSubtaskResult[] {
  const text = raw ?? '';
  const marker = /"subtaskResults"\s*:\s*\[/i.exec(text);
  if (!marker || marker.index === undefined) return [];

  const arrayStart = marker.index + marker[0].length;
  const results: VerifierSubtaskResult[] = [];
  let i = arrayStart;

  while (i < text.length) {
    while (i < text.length && /[\s,]/.test(text[i])) i += 1;
    if (i >= text.length || text[i] === ']') break;
    if (text[i] !== '{') break;

    const obj = extractBalancedObject(text, i);
    if (!obj) break;

    try {
      const parsed = JSON.parse(obj.source) as Partial<VerifierSubtaskResult>;
      if (parsed && typeof parsed.subtaskId === 'string' && parsed.subtaskId.trim()) {
        results.push({
          subtaskId: String(parsed.subtaskId),
          passed: Boolean(parsed.passed),
          feedback: String(parsed.feedback ?? ''),
          issues: Array.isArray(parsed.issues) ? (parsed.issues as VerifierIssue[]) : [],
          inconclusive: false,
        });
      }
    } catch {
      // Truncated object — stop; earlier complete entries remain.
      break;
    }

    i = obj.endIndex;
  }

  return results;
}

function extractBalancedObject(
  text: string,
  start: number,
): { source: string; endIndex: number } | null {
  let depth = 0;
  let inString = false;
  let escape = false;

  for (let i = start; i < text.length; i += 1) {
    const ch = text[i];
    if (inString) {
      if (escape) {
        escape = false;
      } else if (ch === '\\') {
        escape = true;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }

    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') depth += 1;
    if (ch === '}') {
      depth -= 1;
      if (depth === 0) {
        return { source: text.slice(start, i + 1), endIndex: i + 1 };
      }
    }
  }

  return null;
}

/** True when nextStep is an honest executor block (must not be overridden). */
export function isExecutorBlockedSignal(nextStep: string | undefined | null): boolean {
  return Boolean(nextStep && /^\s*blocked\b/i.test(nextStep));
}
