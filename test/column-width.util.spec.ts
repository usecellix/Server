import {
  characterWidthToPoints,
  MIN_READABLE_COLUMN_WIDTH_PT,
  resolveColumnWidthToPoints,
} from '../src/excel-ai/utils/column-width.util';
import { routeShortcutAction } from '../src/excel-ai/utils/shortcut-router.util';

/**
 * TASKS.md #273 — one definition of what SET_COLUMN_WIDTH's `width` means.
 *
 * #265's first fix was a FLOOR (anything under 40 -> 40) and it was too blunt:
 * a live run emitted a sub-40 value for all THIRTEEN columns, so every one
 * floored to exactly 40 and the sheet came out uniformly cramped. Flattening
 * throws away the relative sizing the model got right.
 */
describe('column width resolution (TASKS.md #273)', () => {
  describe('characterWidthToPoints', () => {
    it("round-trips Excel's documented default: 8.43 chars is ~48pt", () => {
      expect(characterWidthToPoints(8.43)).toBe(48);
    });

    it('scales linearly the way Excel does (chars * 7px + 5, then * 0.75)', () => {
      expect(characterWidthToPoints(12)).toBe(67);
      expect(characterWidthToPoints(22)).toBe(119);
    });
  });

  describe('resolveColumnWidthToPoints', () => {
    it('converts the character-count values the model actually emits', () => {
      // The exact widths from the live 13-column run, every one of which the
      // old floor flattened to 40.
      expect(resolveColumnWidthToPoints(8)).toBe(46);
      expect(resolveColumnWidthToPoints(14)).toBe(77);
      expect(resolveColumnWidthToPoints(22)).toBe(119);
    });

    it('PRESERVES relative sizing — the whole point of converting over flooring', () => {
      const unitNo = resolveColumnWidthToPoints(8);
      const guestName = resolveColumnWidthToPoints(22);
      expect(guestName).toBeGreaterThan(unitNo);
      expect(unitNo).not.toBe(guestName);
    });

    it('takes an already-points value at its word', () => {
      for (const pt of [40, 60, 90, 130, 180]) {
        expect(resolveColumnWidthToPoints(pt)).toBe(pt);
      }
    });

    it('still floors a degenerate count that converts below readability', () => {
      expect(resolveColumnWidthToPoints(2)).toBe(MIN_READABLE_COLUMN_WIDTH_PT);
      expect(resolveColumnWidthToPoints(1)).toBe(MIN_READABLE_COLUMN_WIDTH_PT);
    });

    it('leaves non-positive and non-finite values alone (0 belongs to the hide path)', () => {
      expect(resolveColumnWidthToPoints(0)).toBe(0);
      expect(resolveColumnWidthToPoints(-5)).toBe(-5);
      expect(Number.isNaN(resolveColumnWidthToPoints(Number.NaN))).toBe(true);
    });
  });

  /**
   * The same bug from the USER's end: Excel's own UI reports width in
   * characters ("Width: 8.43 (64 pixels)"), so "20 wide" means 20 characters.
   * This fast lane bypasses the Executor normalizer entirely, so before this
   * it produced a 20pt sliver for a perfectly reasonable request.
   */
  describe('the shortcut lane applies the same resolution', () => {
    it('reads "make column B 20 wide" as characters, not points', () => {
      const actions = routeShortcutAction('Set column width of column B to 20', 'Sheet1') ?? [];
      const width = actions.find((a) => a.type === 'SET_COLUMN_WIDTH');
      expect(width).toBeDefined();
      expect(width!.width).toBe(characterWidthToPoints(20));
      expect(width!.width).toBeGreaterThan(MIN_READABLE_COLUMN_WIDTH_PT);
    });

    it('leaves a width already given in points untouched', () => {
      const actions = routeShortcutAction('Set column width of column B to 120', 'Sheet1') ?? [];
      const width = actions.find((a) => a.type === 'SET_COLUMN_WIDTH');
      expect(width!.width).toBe(120);
    });
  });
});
