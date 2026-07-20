/**
 * Extract A1 / Sheet!A1:B2 references from Excel formulas for workbook citations.
 */
const CELL_OR_RANGE =
  /(?:(?:'([^']+)'|([A-Za-z0-9_ ]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/gi;

export function extractFormulaPrecedents(
  formula: string,
  defaultSheet?: string,
): string[] {
  if (!formula || typeof formula !== 'string') return [];
  const refs = new Set<string>();
  const text = formula.startsWith('=') ? formula.slice(1) : formula;

  for (const match of text.matchAll(CELL_OR_RANGE)) {
    const sheet = (match[1] || match[2] || defaultSheet || '').trim();
    const start = `${match[3]}${match[4]}`;
    const end = match[5] && match[6] ? `${match[5]}${match[6]}` : null;
    const address = end ? `${start}:${end}` : start;
    refs.add(sheet ? `${sheet}!${address}` : address);
  }

  return Array.from(refs);
}
