/**
 * TASKS.md #84 — distinguish STRUCTURAL rows from DATA rows.
 *
 * The planner is told each sheet's `rowCount`, but that number counts the header
 * row and every pre-provisioned template row (a row carrying only a seeded
 * formula that currently evaluates to blank). A month sheet built as "headers +
 * 120 empty rows each holding `=(F-E)*G`" reports `rowCount: 121` — which reads
 * to a model as "121 rows of bookings" when the true answer is zero.
 *
 * That ambiguity is what let a request whose premise was false ("some of my bank
 * accounts got renamed — the old ones are still in the sheets", against a
 * workbook with no bookings at all) get planned against as though the data
 * existed. See COMPETITIVE_STUDY_SHORTCUT.md, Trial 2.
 *
 * This module answers the narrower, checkable question: how many rows hold real
 * entered values? It is deliberately conservative — a row counts as data if ANY
 * non-header cell has a non-blank value.
 */

/** A cell is "blank" if it is null/undefined or an empty/whitespace-only string. */
function isBlankCell(value: unknown): boolean {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string') return value.trim() === '';
  return false;
}

/**
 * Count rows below the header that contain at least one non-blank value.
 *
 * `values` is the sheet's loaded cell values. Note these are *computed* values,
 * so a template row whose only content is a formula returning "" correctly
 * counts as empty — which is exactly the case that matters here.
 */
export function countDataRows(values: unknown[][], headerRowIndex: number): number {
  if (!Array.isArray(values) || values.length === 0) return 0;
  const firstDataRow = Math.max(0, headerRowIndex + 1);
  let count = 0;
  for (let r = firstDataRow; r < values.length; r += 1) {
    const row = values[r];
    if (!Array.isArray(row)) continue;
    if (row.some((cell) => !isBlankCell(cell))) count += 1;
  }
  return count;
}

/**
 * One-line, planner-facing description of a sheet's data state.
 *
 * Only emitted when it tells the planner something `rowCount` does not, i.e.
 * when the sheet has structure but no actual data. Staying silent on populated
 * sheets keeps this from becoming noise on every request.
 */
export function describeEmptySheet(
  sheetName: string,
  dataRowCount: number,
  rowCount: number,
): string | null {
  if (dataRowCount > 0) return null;
  const scaffold = rowCount > 1;
  return scaffold
    ? `"${sheetName}": 0 data rows (has headers/template rows only — its ${rowCount} rows are structure, not entered data)`
    : `"${sheetName}": 0 data rows (empty sheet)`;
}
