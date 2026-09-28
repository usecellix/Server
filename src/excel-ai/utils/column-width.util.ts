/**
 * The one definition of what SET_COLUMN_WIDTH's `width` means — TASKS.md #273.
 *
 * The client writes it into Office.js's `Range.format.columnWidth`, which is
 * in POINTS. Almost nobody thinks in points: Excel's own UI reports column
 * width in CHARACTERS ("Width: 8.43 (64 pixels)"), so both the model and a
 * user typing "make column B 20 wide" reach for that scale instead.
 *
 * Two live failures came from this, and they are the same bug from opposite
 * ends. First (#265) the Executor emitted 12-22 for a dashboard, which as
 * points is under half a default column, and every header clipped to a couple
 * of characters. The fix then was a FLOOR — raise anything under 40 to 40 —
 * and that turned out to be too blunt (#273): a later run emitted a sub-40
 * value for all THIRTEEN columns, so every one floored to exactly 40 and the
 * sheet came out uniformly cramped. Flattening throws away the one thing the
 * model got right, the RELATIVE sizing of a wide "Guest Name" against a
 * narrow "Unit No".
 *
 * So convert instead of clamping. Excel maps its character unit to pixels as
 * `chars * MaxDigitWidth + 5` (MaxDigitWidth = 7px for the default Calibri
 * 11) and pixels to points as `px * 0.75`, which round-trips the documented
 * default exactly: 8.43 chars -> 64px -> 48pt.
 */

/** Below this many POINTS a column cannot show real text — half a default column. */
export const MIN_READABLE_COLUMN_WIDTH_PT = 40;

/** Excel's MaxDigitWidth for the default Calibri 11 font, in pixels. */
const MAX_DIGIT_WIDTH_PX = 7;
const PIXELS_TO_POINTS = 0.75;

/** Excel character-width units -> points, via Excel's own pixel formula. */
export function characterWidthToPoints(chars: number): number {
  return Math.round((chars * MAX_DIGIT_WIDTH_PX + 5) * PIXELS_TO_POINTS);
}

/**
 * Resolve a `width` of unknown unit into points.
 *
 * The unit is decided by the readability boundary, and for the same reason it
 * exists: a value under ~40 POINTS is narrower than half a default column and
 * unusable for real text, so it is far more likely a character count than a
 * deliberate choice. At or above it the number is taken at its word.
 *
 * A deliberate sub-40pt spacer column is the one case this gets wrong — it
 * would be widened — but the previous behaviour got that case wrong too (it
 * flattened it to 40), and the common case is now right rather than merely
 * legible. Non-positive widths are left alone: 0 is handled by the hide-column
 * path, not here.
 */
export function resolveColumnWidthToPoints(width: number): number {
  if (!Number.isFinite(width) || width <= 0) return width;
  if (width >= MIN_READABLE_COLUMN_WIDTH_PT) return width;
  // Converted, then floored — a tiny count ("width": 2) still lands under the
  // readable minimum on its own.
  return Math.max(characterWidthToPoints(width), MIN_READABLE_COLUMN_WIDTH_PT);
}
