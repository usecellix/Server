/**
 * Single source of truth for "which row holds the column headers."
 *
 * Previously this heuristic was implemented three times (sheet-analyzer.service.ts,
 * column-slicer.util.ts, and the frontend's sheetAnalyzer.ts) with slightly
 * different, all-broken formulas. Every version accepted a row as the header row
 * once enough of its POPULATED cells looked like text — with no check on how
 * MANY cells were populated. A single merged title cell like
 * "ABC Corp — Purchase Register FY24" is one non-empty cell that is 100% text,
 * so it satisfied the rule instantly and was picked as the header row before the
 * scan ever reached the real headers underneath it. This is exactly what
 * misclassifies sheets whose headers start a few rows down.
 *
 * The fix: a row only qualifies if it is BOTH mostly text AND uses a meaningful
 * share of the sheet's width — a lone title cell can no longer look like a
 * header row just because the one cell it has happens to be text.
 */

const SCAN_WINDOW = 8;
const TEXT_RATIO_THRESHOLD = 0.6;
/** A candidate header row must populate at least this fraction of the sheet's columns. */
const MIN_WIDTH_RATIO = 0.5;
/** ...or at least this many cells outright, for narrow sheets where the ratio is too strict. */
const MIN_WIDTH_ABSOLUTE = 2;

function isNumericLike(text: string): boolean {
  return !Number.isNaN(Number(text.replace(/[,₹\s]/g, '')));
}

/**
 * 0-based index of the most likely header row within the first SCAN_WINDOW rows,
 * or 0 if nothing qualifies (matches prior behavior: fall back to row 0).
 */
export function findHeaderRowIndex(rows: unknown[][], columnCount: number): number {
  const minWidth = Math.max(MIN_WIDTH_ABSOLUTE, Math.ceil(columnCount * MIN_WIDTH_RATIO));

  for (let rowIndex = 0; rowIndex < Math.min(rows.length, SCAN_WINDOW); rowIndex += 1) {
    const row = rows[rowIndex];
    if (!Array.isArray(row)) continue;

    const nonEmpty = row.filter(
      (cell) => cell !== null && cell !== undefined && String(cell).trim() !== '',
    );
    if (nonEmpty.length === 0 || nonEmpty.length < minWidth) continue;

    const textLike = nonEmpty.filter((cell) => {
      const text = String(cell).trim();
      return text.length > 0 && isNumericLike(text) === false;
    });

    if (textLike.length / nonEmpty.length >= TEXT_RATIO_THRESHOLD) {
      return rowIndex;
    }
  }

  return 0;
}
