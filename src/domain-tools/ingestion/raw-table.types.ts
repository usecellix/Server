/**
 * The decoded-file contract between the task pane and the ingestion parsers
 * (Root/ATTACHMENT_EXTRACTION_PLAN.md §4).
 *
 * The task pane turns file bytes into this shape and sends it as JSON, so the
 * raw file (and any PDF password) never leaves the user's machine. It carries
 * no business logic: finding the header, mapping columns and parsing values all
 * happen here on the server, where they are covered by fixtures.
 *
 * Mirrored by `client/src/types/rawTable.ts` — keep the two in step.
 */

export type RawTableSource = 'pdf' | 'xlsx' | 'xls' | 'csv';

export interface RawTableCell {
  /** Cell text exactly as read from the file. */
  t: string;
  /** Left edge in page points. Positioned (PDF) rows only. */
  x0?: number;
  /** Right edge in page points. Positioned (PDF) rows only. */
  x1?: number;
}

export interface RawTableRow {
  /** Where the row came from, shown to the user: "p3 l12" for a PDF, "row 45" for a sheet. */
  ref: string;
  /** 1-based page number. Positioned rows only. */
  page?: number;
  /** Baseline height on the page in points, larger is higher up. Positioned rows only. */
  y?: number;
  cells: RawTableCell[];
}

export interface RawTable {
  source: RawTableSource;
  fileName: string;
  /**
   * `grid`: a cell's index in `cells` is its column (Excel, CSV).
   * `positioned`: cells are text fragments in left-to-right order and columns
   * have to be worked out from `x0`/`x1` (PDF).
   */
  layout: 'grid' | 'positioned';
  rows: RawTableRow[];
}
