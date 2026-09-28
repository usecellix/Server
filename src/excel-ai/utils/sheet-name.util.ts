/** Excel worksheet names: max 31 chars; cannot contain \ / ? * [ ] : */
const INVALID_SHEET_CHARS = /[\\/?*[\]:]/g;
export const MAX_EXCEL_SHEET_NAME_LENGTH = 31;

export const SHEET_NAME_BOUNDARY =
  '(?:\\s+and\\b|\\s+then\\b|\\s+with\\b|\\s+give\\b|\\s+to\\b|\\s*,|\\s*$)';

const UNQUOTED_CALLED_OR_NAMED = new RegExp(
  `(?:named|called)\\s+([A-Za-z][A-Za-z0-9 _-]{0,30}?)(?=${SHEET_NAME_BOUNDARY})`,
  'i',
);

/**
 * Replace every Excel-illegal sheet-name character with a space and collapse
 * whitespace — no length truncation, no fallback for an empty result. Use this
 * (instead of `sanitizeExcelSheetName`) when the caller still needs to reserve
 * room in the 31-char budget for something appended afterward (e.g. a layout
 * suffix like " — Books Only") before truncating the combined name.
 */
export function stripIllegalSheetNameChars(raw: string): string {
  return raw.trim().replace(INVALID_SHEET_CHARS, ' ').replace(/\s+/g, ' ').trim();
}

export function sanitizeExcelSheetName(raw: string, fallback = 'Sheet'): string {
  let name = stripIllegalSheetNameChars(raw);
  if (!name) name = fallback;
  if (name.length > MAX_EXCEL_SHEET_NAME_LENGTH) {
    name = name.slice(0, MAX_EXCEL_SHEET_NAME_LENGTH).trim();
  }
  return name || fallback;
}

export function extractSheetNameFromPrompt(message: string): string | null {
  const quoted = /(?:named|called)\s+["']([^"']+)["']/i.exec(message)?.[1];
  if (quoted?.trim()) return sanitizeExcelSheetName(quoted.trim());

  const unquoted = UNQUOTED_CALLED_OR_NAMED.exec(message)?.[1];
  if (unquoted?.trim()) return sanitizeExcelSheetName(unquoted.trim());

  return null;
}

export function detectCompoundSheetFollowUp(message: string): boolean {
  return (
    /\s+and\s+(?:give|add|create|build|make|show|sort|fill|chart|analy(?:sis|ze|se)|insert|populate|copy|duplicate)\b/i.test(
      message,
    ) ||
    /\s+then\s+(?:give|add|create|build|make|show|sort|fill|chart|analy(?:sis|ze|se))\b/i.test(
      message,
    )
  );
}
