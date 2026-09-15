import { SheetActionPayload } from '../types/sheet-actions.types';
import { columnIndexToLetter } from './consolidation-pass.util';

/**
 * Styling for a "categorized reconciliation" output sheet (navy section/header bands,
 * white bold text, sensible column widths, currency/date number formats, alternating
 * row banding, frozen title row) — built once here so every recon-type sheet (purchase
 * or sales) that uses this shape inherits the same look, instead of duplicating format
 * logic per call site.
 */

const NAVY_FILL = '#203764';
const BAND_FILL = '#F2F2F2';
const WHITE = '#FFFFFF';
const CURRENCY_NUMBER_FORMAT = '#,##0.00';
const DATE_NUMBER_FORMAT = 'dd-mmm-yyyy';

const CURRENCY_HEADER_RE =
  /taxable value|\bigst\b|\bcgst\b|\bsgst\b|tax amount|itc amount|invoice value/i;
const DATE_HEADER_RE = /invoice date/i;
const WIDE_HEADER_RE = /vendor name|explanation|particulars|narration/i;
const NARROW_HEADER_RE = /^gstin|gstin$|source row|^reason$|document type|^pass$|^status$|^rcm$|^confidence$/i;

const WIDE_COLUMN_WIDTH = 320; // ~45 characters
const NARROW_COLUMN_WIDTH = 110; // ~16 characters
const DEFAULT_COLUMN_WIDTH = 140; // ~20 characters

export interface StyledSheetSection {
  /** 0-based row index of this section's title/label row (e.g. "Blank GSTIN rows"). */
  titleRow: number;
  /** 0-based row index of the column-header row immediately under the title. */
  headerRow: number;
  /** Column headers for this section, in order — used to infer currency/date columns and widths. */
  headers: string[];
  /** 0-based inclusive [first, last] data row range, or null when the section has no data rows. */
  dataRowRange: [number, number] | null;
}

export interface StyledSheetPlan {
  sheetName: string;
  /** Total column count every row in this sheet is padded to. */
  columnCount: number;
  /** 0-based row index of the overall sheet title (typically row 0). */
  sheetTitleRow: number;
  sections: StyledSheetSection[];
  /**
   * 0-based inclusive [first, last] row range of the SUMMARY block's "label, value"
   * metric rows (e.g. "Matched (exact)" / "Blank GSTIN in books" / ...) — excludes the
   * "SUMMARY" title row itself and the blank spacer rows around it. The value always
   * sits in column B (index 1), per the builder's `pushRow([label, value])` convention.
   */
  summaryValueRange?: [number, number] | null;
}

function rangeRef(sheetName: string, colStart: number, colEnd: number, rowStart: number, rowEnd: number) {
  const startLetter = columnIndexToLetter(colStart);
  const endLetter = columnIndexToLetter(colEnd);
  return `${startLetter}${rowStart + 1}:${endLetter}${rowEnd + 1}`;
}

function bandFormatAction(plan: StyledSheetPlan, dataRowRange: [number, number]): SheetActionPayload {
  return {
    type: 'CONDITIONAL_FORMAT',
    sheetName: plan.sheetName,
    range: rangeRef(plan.sheetName, 0, plan.columnCount - 1, dataRowRange[0], dataRowRange[1]),
    rule: { kind: 'formula', formula: '=MOD(ROW(),2)=0', format: { fillColor: BAND_FILL } },
  };
}

function navyHeaderAction(
  plan: StyledSheetPlan,
  row: number,
  opts: { center?: boolean } = {},
): SheetActionPayload {
  return {
    type: 'FORMAT_RANGE',
    sheetName: plan.sheetName,
    range: rangeRef(plan.sheetName, 0, plan.columnCount - 1, row, row),
    format: {
      bold: true,
      fontColor: WHITE,
      fillColor: NAVY_FILL,
      ...(opts.center ? { horizontalAlignment: 'center' as const } : {}),
    },
  };
}

/**
 * Builds the FORMAT_RANGE / SET_COLUMN_WIDTH / FREEZE_PANES / CONDITIONAL_FORMAT actions
 * for a styled recon sheet. Append the result AFTER the CREATE_SHEET + WRITE_TABLE
 * actions — WRITE_TABLE always writes at A1, so row/column indices in `plan` translate
 * directly to sheet rows/columns.
 */
export function buildReconSheetStylingActions(plan: StyledSheetPlan): SheetActionPayload[] {
  const actions: SheetActionPayload[] = [];

  // 1. Overall sheet title row.
  actions.push(navyHeaderAction(plan, plan.sheetTitleRow));

  // 2. Per-section title + column-header rows, number formats, and row banding.
  for (const section of plan.sections) {
    actions.push(navyHeaderAction(plan, section.titleRow));
    actions.push(navyHeaderAction(plan, section.headerRow, { center: true }));

    if (section.dataRowRange) {
      const [start, end] = section.dataRowRange;
      section.headers.forEach((header, colIdx) => {
        let numberFormat: string | null = null;
        if (CURRENCY_HEADER_RE.test(header)) numberFormat = CURRENCY_NUMBER_FORMAT;
        else if (DATE_HEADER_RE.test(header)) numberFormat = DATE_NUMBER_FORMAT;
        if (!numberFormat) return;
        actions.push({
          type: 'FORMAT_RANGE',
          sheetName: plan.sheetName,
          range: rangeRef(plan.sheetName, colIdx, colIdx, start, end),
          format: { numberFormat },
        });
      });
      // Center every column, including the long-text ones (Vendor Name, Explanation) —
      // alignment only, no width/height/wrap change, so anything that already overflows
      // or clips keeps doing so, just centered instead of Excel's left/right default.
      actions.push({
        type: 'FORMAT_RANGE',
        sheetName: plan.sheetName,
        range: rangeRef(plan.sheetName, 0, plan.columnCount - 1, start, end),
        format: { horizontalAlignment: 'center' },
      });
      actions.push(bandFormatAction(plan, section.dataRowRange));
    }
  }

  // 2b. SUMMARY block's value column (B) — same alignment-only treatment as the detail
  //     sections: center, no numberFormat/wrapText/width/height change.
  if (plan.summaryValueRange) {
    const [start, end] = plan.summaryValueRange;
    actions.push({
      type: 'FORMAT_RANGE',
      sheetName: plan.sheetName,
      range: rangeRef(plan.sheetName, 1, 1, start, end),
      format: { horizontalAlignment: 'center' },
    });
  }

  // 3. Column widths — widest "wide-like" or narrowest "narrow-like" classification seen
  //    for that column index across every section (sections share the sheet's column grid).
  const widthByCol = new Map<number, number>();
  for (const section of plan.sections) {
    section.headers.forEach((header, colIdx) => {
      const width = WIDE_HEADER_RE.test(header)
        ? WIDE_COLUMN_WIDTH
        : NARROW_HEADER_RE.test(header)
          ? NARROW_COLUMN_WIDTH
          : DEFAULT_COLUMN_WIDTH;
      const existing = widthByCol.get(colIdx);
      if (existing === undefined || width > existing) widthByCol.set(colIdx, width);
    });
  }
  for (let col = 0; col < plan.columnCount; col++) {
    actions.push({
      type: 'SET_COLUMN_WIDTH',
      sheetName: plan.sheetName,
      col,
      width: widthByCol.get(col) ?? DEFAULT_COLUMN_WIDTH,
    });
  }

  // 4. Freeze the sheet title row so it stays visible while scrolling through any section.
  actions.push({
    type: 'FREEZE_PANES',
    sheetName: plan.sheetName,
    freezeRows: plan.sheetTitleRow + 1,
  });

  return actions;
}

/** Small incremental builder so callers can push rows and record section boundaries in one pass, instead of computing row-index arithmetic by hand. */
export class StyledSheetBuilder {
  readonly rows: unknown[][] = [];
  readonly sections: StyledSheetSection[] = [];
  /** Set by beginSummaryBlock()/endSummaryBlock() — see StyledSheetPlan.summaryValueRange. */
  summaryValueRange: [number, number] | null = null;
  private readonly width: number;
  private summaryStart: number | null = null;

  constructor(width: number) {
    this.width = width;
  }

  /** Call immediately before the first "label, value" SUMMARY metric row. */
  beginSummaryBlock(): void {
    this.summaryStart = this.rows.length;
  }

  /** Call immediately after the last "label, value" SUMMARY metric row. */
  endSummaryBlock(): void {
    if (this.summaryStart === null) return;
    this.summaryValueRange = [this.summaryStart, this.rows.length - 1];
    this.summaryStart = null;
  }

  private padRow(row: unknown[]): unknown[] {
    if (row.length >= this.width) return row;
    return [...row, ...new Array(this.width - row.length).fill('')];
  }

  /** Push a single row (padded to sheet width) with no section tracking — for title/summary/spacer rows. */
  pushRow(row: unknown[]): number {
    this.rows.push(this.padRow(row));
    return this.rows.length - 1;
  }

  /**
   * Push a labeled section: title row, column-header row, each data row, then a blank
   * spacer — and record its boundaries for styling. No-ops (and records nothing) when
   * `dataRows` is empty, matching the existing "omit empty sections" sheet behavior.
   */
  pushSection(title: string, headers: string[], dataRows: unknown[][]): void {
    if (!dataRows.length) return;
    const titleRow = this.pushRow([title]);
    const headerRow = this.pushRow(headers);
    const dataStart = this.rows.length;
    for (const row of dataRows) this.pushRow(row);
    const dataEnd = this.rows.length - 1;
    this.pushRow(['']);
    this.sections.push({ titleRow, headerRow, headers, dataRowRange: [dataStart, dataEnd] });
  }
}
