import { Action, SubTask } from '../types/agent.types';
import { formulaSheetRef } from '../../excel-ai/utils/consolidation-pass.util';

/**
 * The Main dashboard of a repeated-sheet ledger, built by code — TASKS.md #327.
 *
 * Every live run of the 12-month ledger this week built Main with the model,
 * and every run got Main wrong a DIFFERENT way while reporting "All steps
 * applied": `operations: [24]` instead of July–December (#322), a KPI band
 * summing B5:B10 when told B5:B16, a title never written, a "Pending" SUMIF
 * over a status list that has no "Pending", a chart plotting "Month" as a
 * series, and a layout the user rejected ("it needs to be like a dashboard,
 * proper spacing, all content"). None of it is a judgement call: once the
 * month sheets' header row is known (Phase 1.5 resolves it), where Total
 * Amount and Payment Status live is known, and so is every formula Main needs.
 * The same reasoning that made the month headers deterministic (#280) makes
 * the dashboard deterministic: it cannot drift, elide, or mis-sum.
 *
 * Layout (one row per month plus a grand total — the shape the user chose,
 * TASKS.md #324):
 *
 *   1  Payments Dashboard
 *   2  (how it updates)
 *   4        Total Amount   Paid        Pending     Bookings      <- KPI labels
 *   5        ₹ …            ₹ …         ₹ …         n             <- = the Total row
 *   7  Monthly Summary
 *   8  Month | Total Amount | Paid | Pending | Bookings           <- header
 *   9  January … December                                         <- one row each
 *   21 Total                                                      <- grand total
 *   chart of rows 8–20 (months only) to the right, from column G
 *   23 All Bookings
 *   24 Month | <every column of the month sheets>                    <- header
 *   25 one live row per booking, from every month                     <- see below
 *
 * The All Bookings rows are the user's "all the details of the remaining
 * sheets" (TASKS.md #333 — the same user, seeing the per-month-only Main, said
 * the prompt asks for all details and "below it is blank"). This step writes
 * only the section's header row; the consolidation pass (#142/#269/#312)
 * recognises "Month + the month sheets' exact header row" and adds the single
 * spilling formula under it at run time, where it knows whether this Excel
 * supports dynamic arrays and can fall back to an honest note if not.
 */

/** Sheet names the planner uses for the dashboard tab. */
const DASHBOARD_SHEET = /^(main|dashboard|summary|overview)$/i;
/** Smallest repeated-sheet group worth a dashboard. */
const MIN_SOURCE_SHEETS = 3;
const CURRENCY_FORMAT = '₹ #,##,##0.00';
const TITLE_ROW = 1;
const NOTE_ROW = 2;
const KPI_LABEL_ROW = 4;
const KPI_VALUE_ROW = 5;
const SECTION_ROW = 7;
const HEADER_ROW = 8;
const FIRST_MONTH_ROW = 9;
/** Rows formatted under the All Bookings header — the formula spills into them. */
const BOOKINGS_FORMAT_ROWS = 1000;
const DATE_HEADER = /date|check ?in|check ?out/i;
const MONEY_HEADER = /amount|rate|total|balance|received|price|paid|due/i;
const HEADER_FILL = '#1F4E78';
const TOTAL_FILL = '#F2F2F2';
const MUTED = '#6B7280';

export interface DashboardShape {
  dashboardSheet: string;
  sourceSheets: string[];
  /** The deterministic header steps that create the source sheets. */
  sourceStepIds: string[];
  totalColumn: string;
  statusColumn?: string;
  /** The source sheets' full header row, for the All Bookings section. */
  sourceHeaders: string[];
}

function letter(oneBased: number): string {
  let n = oneBased;
  let out = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    out = String.fromCharCode(65 + mod) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

const norm = (label: string) => label.trim().toLowerCase().replace(/\s+/g, ' ');

function findColumn(headers: string[], test: (label: string) => boolean): string | undefined {
  const index = headers.findIndex((h) => test(norm(h)));
  return index >= 0 ? letter(index + 1) : undefined;
}

/**
 * The ledger shape this applies to: at least three deterministic header steps
 * sharing one header row that has a Total Amount column, plus planned work on
 * a dashboard-named sheet. Anything else returns null and the plan is left to
 * the model exactly as before.
 */
export function findDashboardShape(subtasks: SubTask[]): DashboardShape | null {
  const groups = new Map<string, SubTask[]>();
  for (const step of subtasks) {
    if (!step.isDeterministicHeaderStep || !step.resolvedHeaderRow?.length) continue;
    const key = step.resolvedHeaderRow.map(norm).join('|');
    groups.set(key, [...(groups.get(key) ?? []), step]);
  }
  const sources = [...groups.values()].sort((a, b) => b.length - a.length)[0];
  if (!sources || sources.length < MIN_SOURCE_SHEETS) return null;

  const headers = sources[0].resolvedHeaderRow!;
  const totalColumn = findColumn(headers, (h) => h === 'total amount' || h === 'total');
  if (!totalColumn) return null;
  const statusColumn = findColumn(headers, (h) => h === 'payment status' || h === 'status');

  const sourceNames = new Set(sources.map((s) => norm(s.targetSheet)));
  const dashboard = subtasks.find(
    (s) => DASHBOARD_SHEET.test(s.targetSheet.trim()) && !sourceNames.has(norm(s.targetSheet)),
  );
  if (!dashboard) return null;

  return {
    dashboardSheet: dashboard.targetSheet.trim(),
    sourceSheets: sources.map((s) => s.targetSheet),
    sourceStepIds: sources.map((s) => s.id),
    totalColumn,
    statusColumn,
    sourceHeaders: headers,
  };
}

interface Column {
  header: string;
  kpiLabel: string;
  /** Formula for the given month sheet on the given dashboard row. */
  perSheet: (ref: string, row: number, self: Record<string, string>) => string;
  numberFormat: string;
}

export function buildDashboardActions(shape: DashboardShape): Action[] {
  const sheet = shape.dashboardSheet;
  const t = shape.totalColumn;
  const s = shape.statusColumn;

  // Paid/Pending only when there is a status column to classify by. Pending is
  // what is not yet Paid — NOT a SUMIF on the word "Pending", which the live
  // status list (Paid / Partial / Unpaid) never contains.
  const columns: Column[] = [
    {
      header: 'Total Amount',
      kpiLabel: 'Total Amount',
      perSheet: (ref) => `=SUM(${ref}!${t}:${t})`,
      numberFormat: CURRENCY_FORMAT,
    },
    ...(s
      ? [
          {
            header: 'Paid',
            kpiLabel: 'Paid',
            perSheet: (ref: string) => `=SUMIF(${ref}!${s}:${s},"Paid",${ref}!${t}:${t})`,
            numberFormat: CURRENCY_FORMAT,
          },
          {
            header: 'Pending',
            kpiLabel: 'Pending',
            perSheet: (_ref: string, row: number, self: Record<string, string>) =>
              `=${self['Total Amount']}${row}-${self.Paid}${row}`,
            numberFormat: CURRENCY_FORMAT,
          },
        ]
      : []),
    {
      header: 'Bookings',
      kpiLabel: 'Bookings',
      perSheet: (ref) => `=COUNT(${ref}!${t}:${t})`,
      numberFormat: '0',
    },
  ];

  const colOf: Record<string, string> = {};
  columns.forEach((c, i) => (colOf[c.header] = letter(i + 2)));
  const lastCol = columns.length + 1; // column A is the month label
  const lastMonthRow = FIRST_MONTH_ROW + shape.sourceSheets.length - 1;
  const totalRow = lastMonthRow + 1;

  const ops: Array<{ address: string; value?: string; formula?: string }> = [
    { address: `A${TITLE_ROW}`, value: 'Payments Dashboard' },
    {
      address: `A${NOTE_ROW}`,
      value: `Totals update automatically from the ${shape.sourceSheets.length} month sheets.`,
    },
    { address: `A${SECTION_ROW}`, value: 'Monthly Summary' },
    { address: `A${HEADER_ROW}`, value: 'Month' },
    { address: `A${totalRow}`, value: 'Total' },
  ];
  columns.forEach((c) => {
    const col = colOf[c.header];
    ops.push({ address: `${col}${KPI_LABEL_ROW}`, value: c.kpiLabel });
    ops.push({ address: `${col}${KPI_VALUE_ROW}`, formula: `=${col}${totalRow}` });
    ops.push({ address: `${col}${HEADER_ROW}`, value: c.header });
    ops.push({ address: `${col}${totalRow}`, formula: `=SUM(${col}${FIRST_MONTH_ROW}:${col}${lastMonthRow})` });
  });
  shape.sourceSheets.forEach((name, i) => {
    const row = FIRST_MONTH_ROW + i;
    const ref = formulaSheetRef(name);
    ops.push({ address: `A${row}`, value: name });
    columns.forEach((c) => ops.push({ address: `${colOf[c.header]}${row}`, formula: c.perSheet(ref, row, colOf) }));
  });

  const fmt = (row: number, col: number, rowCount: number, colCount: number, format: object): Action =>
    ({ type: 'FORMAT_RANGE', sheetName: sheet, row: row - 1, col, rowCount, colCount, format }) as Action;

  const numberFormats: Action[] = columns.map((c, i) =>
    fmt(FIRST_MONTH_ROW, i + 1, totalRow - FIRST_MONTH_ROW + 1, 1, { numberFormat: c.numberFormat }),
  );
  const kpiFormats: Action[] = columns.map((c, i) =>
    fmt(KPI_VALUE_ROW, i + 1, 1, 1, { numberFormat: c.numberFormat, bold: true, fontSize: 14 }),
  );

  // Months only (header + rows), never the Total row — it would dwarf every bar.
  const chartColumns = columns.filter((c) => c.numberFormat === CURRENCY_FORMAT).length;
  const chartEndRow = totalRow + 1;
  const bookingsTitleRow = totalRow + 2;
  const bookingsHeaderRow = bookingsTitleRow + 1;
  const bookingsHeaders = ['Month', ...shape.sourceHeaders];
  ops.push({ address: `A${bookingsTitleRow}`, value: 'All Bookings' });
  bookingsHeaders.forEach((label, i) => ops.push({ address: `${letter(i + 1)}${bookingsHeaderRow}`, value: label }));
  const bookingsFormats: Action[] = shape.sourceHeaders.flatMap((label, i) => {
    const numberFormat = DATE_HEADER.test(label)
      ? 'dd-mm-yyyy'
      : MONEY_HEADER.test(label)
        ? CURRENCY_FORMAT
        : undefined;
    return numberFormat ? [fmt(bookingsHeaderRow + 1, i + 1, BOOKINGS_FORMAT_ROWS, 1, { numberFormat })] : [];
  });
  const widestCol = Math.max(lastCol, bookingsHeaders.length);

  return [
    // `position` is numeric for ADD_SHEET on the wire; the payload union types it for columns.
    { type: 'ADD_SHEET', name: sheet, sheetName: sheet, position: 0 } as unknown as Action,
    { type: 'BATCH_SET', sheetName: sheet, operations: ops } as Action,
    fmt(TITLE_ROW, 0, 1, 1, { bold: true, fontSize: 16 }),
    fmt(NOTE_ROW, 0, 1, 1, { italic: true, fontColor: MUTED }),
    fmt(KPI_LABEL_ROW, 1, 1, columns.length, { fontColor: MUTED, horizontalAlignment: 'right' }),
    ...kpiFormats,
    fmt(SECTION_ROW, 0, 1, 1, { bold: true, fontSize: 12 }),
    fmt(HEADER_ROW, 0, 1, lastCol, { bold: true, fontColor: '#FFFFFF', fillColor: HEADER_FILL }),
    ...numberFormats,
    fmt(totalRow, 0, 1, lastCol, { bold: true, fillColor: TOTAL_FILL, borders: 'outer' }),
    fmt(bookingsTitleRow, 0, 1, 1, { bold: true, fontSize: 12 }),
    fmt(bookingsHeaderRow, 0, 1, bookingsHeaders.length, { bold: true, fontColor: '#FFFFFF', fillColor: HEADER_FILL }),
    ...bookingsFormats,
    { type: 'SET_COLUMN_WIDTH', sheetName: sheet, col: 0, colCount: 1, columns: ['A'], width: 110 } as Action,
    {
      type: 'SET_COLUMN_WIDTH',
      sheetName: sheet,
      col: 1,
      colCount: widestCol - 1,
      columns: Array.from({ length: widestCol - 1 }, (_, i) => letter(i + 2)),
      width: 95,
    } as Action,
    { type: 'HIDE_GRIDLINES', sheetName: sheet } as Action,
    {
      type: 'CREATE_CHART',
      sheetName: sheet,
      sourceSheetName: sheet,
      sourceRange: `A${HEADER_ROW}:${letter(chartColumns + 1)}${lastMonthRow}`,
      chartType: 'ColumnClustered',
      title: 'Monthly Payments',
      startCell: `${letter(lastCol + 2)}${SECTION_ROW}`,
      endCell: `${letter(lastCol + 9)}${chartEndRow}`,
    } as Action,
  ];
}

/**
 * Replace every model-planned subtask on the dashboard sheet with ONE
 * deterministic step. Dependencies on the replaced subtasks are pointed at it.
 */
export function replaceDashboardSubtasks(subtasks: SubTask[]): {
  subtasks: SubTask[];
  replaced: string[];
  shape: DashboardShape | null;
} {
  const shape = findDashboardShape(subtasks);
  if (!shape) return { subtasks, replaced: [], shape: null };

  const onDashboard = (s: SubTask) => norm(s.targetSheet) === norm(shape.dashboardSheet);
  const replaced = subtasks.filter(onDashboard).map((s) => s.id);
  const replacedSet = new Set(replaced);
  const usedIds = new Set(subtasks.map((s) => s.id));
  let id = `dash_${shape.dashboardSheet.replace(/[^A-Za-z0-9]/g, '') || 'Main'}`;
  while (usedIds.has(id)) id = `${id}_2`;

  const dashboardStep: SubTask = {
    id,
    description:
      `Build the '${shape.dashboardSheet}' dashboard — KPI totals, one row per sheet ` +
      `(${shape.sourceSheets.join(', ')}) with ${shape.statusColumn ? 'Total / Paid / Pending / Bookings' : 'Total / Bookings'}, ` +
      `a Total row and a chart (built automatically).`,
    targetSheet: shape.dashboardSheet,
    dependsOn: shape.sourceStepIds,
    estimatedActions: 1,
    deterministicActions: buildDashboardActions(shape),
  };

  const out: SubTask[] = [];
  let inserted = false;
  for (const subtask of subtasks) {
    if (replacedSet.has(subtask.id)) {
      if (!inserted) {
        out.push(dashboardStep);
        inserted = true;
      }
      continue;
    }
    const deps = subtask.dependsOn.map((dep) => (replacedSet.has(dep) ? id : dep));
    out.push({ ...subtask, dependsOn: [...new Set(deps)] });
  }
  return { subtasks: out, replaced, shape };
}
