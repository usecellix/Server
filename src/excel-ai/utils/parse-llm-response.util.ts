import { SheetActionPayload } from '../types/sheet-actions.types';

export interface ParsedLlmPayload {
  type?: string;
  question?: string;
  options?: string[];
  answer?: string;
  actions?: SheetActionPayload[];
  explanation?: string;
}

export function extractJsonFromLlmText(text: string): ParsedLlmPayload | null {
  const trimmed = text.trim();
  if (!trimmed) return null;

  const candidates: string[] = [];

  const fencePattern = /```(?:json)?\s*([\s\S]*?)```/gi;
  let fenceMatch: RegExpExecArray | null;
  while ((fenceMatch = fencePattern.exec(trimmed)) !== null) {
    candidates.push(fenceMatch[1].trim());
  }

  candidates.push(trimmed);

  const braceBlocks = findJsonObjectStrings(trimmed);
  candidates.push(...braceBlocks);

  for (const candidate of candidates) {
    const parsed = tryParse(candidate);
    if (parsed) return parsed;
  }

  return null;
}

function findJsonObjectStrings(text: string): string[] {
  const results: string[] = [];
  let start = text.indexOf('{');
  while (start >= 0) {
    let depth = 0;
    for (let i = start; i < text.length; i += 1) {
      if (text[i] === '{') depth += 1;
      if (text[i] === '}') depth -= 1;
      if (depth === 0) {
        results.push(text.slice(start, i + 1));
        start = text.indexOf('{', i + 1);
        break;
      }
    }
    if (depth !== 0) break;
  }
  return results.reverse();
}

function tryParse(raw: string): ParsedLlmPayload | null {
  try {
    const parsed = JSON.parse(raw) as ParsedLlmPayload;
    if (parsed && typeof parsed === 'object') {
      return parsed;
    }
  } catch {
    // continue
  }
  return null;
}

export function hasActionPayload(parsed: ParsedLlmPayload): boolean {
  return Array.isArray(parsed.actions) && parsed.actions.length > 0;
}
