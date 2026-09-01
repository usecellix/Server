import { WorkbookContext } from '../../types/cellix.types';
import { SheetActionPayload } from '../types/sheet-actions.types';
import { extractSheetMentions, stripSheetMentions } from './sheet-mentions.util';

export function detectDeleteSheetIntent(message: string): boolean {
  const lower = message.toLowerCase();
  return (
    /\b(delete|remove|drop)\b/.test(lower) &&
    (/\bsheets?\b/.test(lower) || /\btab(s)?\b/.test(lower))
  );
}

/**
 * "Delete all sheets except X" / "keep X remove the rest" — named sheets are preserves, not targets.
 */
export function isPreserveOtherSheetsDelete(message: string): boolean {
  const lower = message.toLowerCase();
  if (/\b(except|except for|apart from|other than)\b/.test(lower)) return true;
  if (
    /\b(keep|keeping|leave|leaving)\b/.test(lower) &&
    /\b(all|other|every|rest|remaining)\b/.test(lower)
  ) {
    return true;
  }
  if (
    /\b(delete|remove|drop)\s+(all|every)\b/.test(lower) &&
    /\b(keep|except|but)\b/.test(lower)
  ) {
    return true;
  }
  return false;
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function extractQuotedNames(message: string): string[] {
  const names: string[] = [];
  const pattern = /["']([^"']+)["']/g;
  let match = pattern.exec(message);
  while (match) {
    const name = match[1]?.trim();
    if (name) names.push(name);
    match = pattern.exec(message);
  }
  return names;
}

function resolveSheetNames(candidates: string[], availableSheets: string[]): string[] {
  const byLower = new Map(availableSheets.map((sheet) => [sheet.toLowerCase(), sheet]));
  const resolved: string[] = [];

  for (const candidate of candidates) {
    const trimmed = candidate.trim();
    if (!trimmed) continue;
    const exact = byLower.get(trimmed.toLowerCase());
    if (exact && !resolved.includes(exact)) {
      resolved.push(exact);
    }
  }

  return resolved;
}

export function listWorkbookSheetNames(workbookContext?: WorkbookContext | null): string[] {
  if (!workbookContext?.sheets?.length) return [];
  return workbookContext.sheets
    .map((sheet) => sheet.sheetName)
    .filter((name): name is string => Boolean(name));
}

/** Sheets named in the message (mentions, quotes, active-sheet phrase, or workbook name scan). */
export function extractReferencedSheetNames(
  message: string,
  availableSheets: string[],
  activeSheet?: string,
): string[] {
  const mentions = extractSheetMentions(message);
  if (mentions.length > 0) {
    const resolved = resolveSheetNames(mentions, availableSheets);
    if (resolved.length > 0) return resolved;
    return mentions;
  }

  if (/\b(this|current|active)\s+sheet\b/i.test(message) && activeSheet) {
    return [activeSheet];
  }

  const quoted = extractQuotedNames(message);
  if (quoted.length > 0) {
    const resolved = resolveSheetNames(quoted, availableSheets);
    if (resolved.length > 0) return resolved;
    return quoted;
  }

  const cleaned = stripSheetMentions(message);
  const sortedSheets = [...availableSheets].sort((a, b) => b.length - a.length);
  const mentioned: string[] = [];
  for (const sheet of sortedSheets) {
    const pattern = new RegExp(`\\b${escapeRegex(sheet)}\\b`, 'i');
    if (pattern.test(cleaned) && !mentioned.includes(sheet)) {
      mentioned.push(sheet);
    }
  }
  if (mentioned.length > 0) {
    mentioned.sort(
      (a, b) =>
        cleaned.toLowerCase().indexOf(a.toLowerCase()) -
        cleaned.toLowerCase().indexOf(b.toLowerCase()),
    );
    return mentioned;
  }

  // "except Purchase Register" / "keep Sales" when sheet name is not exact-matched above
  const exceptClause =
    /\b(?:except(?:\s+for)?|apart\s+from|other\s+than|keep(?:ing)?|leave(?:ing)?)\s+(.+?)(?:[.!?]|$)/i.exec(
      cleaned,
    );
  if (exceptClause?.[1]) {
    const part = exceptClause[1]
      .replace(/\b(?:sheet|tab)s?\b/gi, ' ')
      .replace(/\b(?:named|called)\b/gi, ' ')
      .trim();
    const parts = part
      .split(/\s*,\s*|\s+and\s+/i)
      .map((p) => p.replace(/^["']|["']$/g, '').trim())
      .filter(Boolean);
    const resolved = resolveSheetNames(parts, availableSheets);
    if (resolved.length > 0) return resolved;
  }

  return [];
}

export function extractDeleteSheetNames(
  message: string,
  availableSheets: string[],
  activeSheet?: string,
): string[] {
  if (isPreserveOtherSheetsDelete(message)) {
    if (availableSheets.length === 0) return [];
    const preserve = extractReferencedSheetNames(message, availableSheets, activeSheet);
    if (preserve.length === 0) {
      // Ambiguous — do not invent deletes (esp. never treat "except X" as delete X).
      return [];
    }
    const preserveSet = new Set(preserve.map((name) => name.toLowerCase()));
    return availableSheets.filter((name) => !preserveSet.has(name.toLowerCase()));
  }

  const referenced = extractReferencedSheetNames(message, availableSheets, activeSheet);
  if (referenced.length > 0) return referenced;

  const cleaned = stripSheetMentions(message);
  const listMatch =
    /\b(?:delete|remove|drop)\s+(?:the\s+)?sheets?\s+(?:named\s+)?(.+?)(?:[.!?]|$)/i.exec(
      cleaned,
    );
  if (listMatch?.[1]) {
    const parts = listMatch[1]
      .split(/\s*,\s*|\s+and\s+/i)
      .map((part) => part.replace(/^["']|["']$/g, '').trim())
      .filter(Boolean);
    const resolved = resolveSheetNames(parts, availableSheets);
    if (resolved.length > 0) return resolved;
    return parts;
  }

  return [];
}

export function tryLocalDeleteSheetActions(
  message: string,
  workbookContext?: WorkbookContext | null,
): SheetActionPayload[] | null {
  if (!detectDeleteSheetIntent(message)) return null;

  const availableSheets = listWorkbookSheetNames(workbookContext);
  const hasMentions = extractSheetMentions(message).length > 0;
  if (availableSheets.length === 0 && !hasMentions) return null;

  // Preserve-all-except needs workbook list; without it we would mis-delete keep sheet.
  if (isPreserveOtherSheetsDelete(message) && availableSheets.length === 0) {
    return null;
  }

  const sheetNames = extractDeleteSheetNames(
    message,
    availableSheets,
    workbookContext?.activeSheet,
  );
  if (sheetNames.length === 0) return null;

  return sheetNames.map((sheetName) => ({
    type: 'DELETE_SHEET' as const,
    sheetName,
  }));
}

export function buildDeleteSheetAnswer(sheetNames: string[]): string {
  if (sheetNames.length === 1) {
    return `Delete sheet **${sheetNames[0]}**`;
  }
  return `Delete sheets: ${sheetNames.map((name) => `**${name}**`).join(', ')}`;
}
