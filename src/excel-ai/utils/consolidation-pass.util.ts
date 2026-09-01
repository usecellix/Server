import { SheetActionPayload } from '../types/sheet-actions.types';
import { detectHeaderRuns, HeaderRun } from './presentation-pass.util';
import { sheetsCreatedInBatch } from './sheet-header-state.util';

/**
 * Deterministic consolidation pass — TASKS.md #142.
 *
 * When the user asks for "a main sheet [that] has all the details of the
 * remaining sheets", the planner writes the consolidated table's HEADER row and
 * nothing else. That is deliberate — planner rule 3 forbids planning
 * COPY_FILTERED_RANGE against month sheets that were just created empty, and it
 * is right to: there are no rows to copy at plan time, and a copy would be a
 * one-shot snapshot anyway.
 *
 * But the result is a table that never fills in. The reported symptom was
 * exactly that: a booking typed into `March` never appeared on `Main`. A copy
 * would not have fixed it either — the user's expectation is a *live* view.
 *
 * So: one spilling dynamic-array formula, written once, that stacks every month
 * sheet's data range, tags each row with its sheet name, and filters out the
 * empty rows. It updates the moment anything is typed into any month sheet,
 * needs no row-by-row formulas, and cannot go stale.
 *
 * **Requires Excel 365 / 2021+** (`LET`, `VSTACK`, `HSTACK`, `FILTER`, `BYROW`,
 * `LAMBDA`, `EXPAND`, `DROP`). On older Excel the cell shows `#NAME?`. This was
 * an explicit product decision, not an oversight.
 */

/** Rows per source sheet the consolidated view reaches into. */
const SOURCE_ROW_LIMIT = 500;

/** Minimum source sheets sharing a schema before this is a consolidation at all. */
const MIN_SOURCE_SHEETS = 2;

/**
 * Separator for schema signature keys. A control character rather than a space
 * so it can never collide with a real column label containing spaces.
 */
const SCHEMA_KEY_SEPARATOR = String.fromCharCode(31);

/**
 * Smallest consolidated table worth detecting: an origin column plus at least
 * two real data columns. Below that the schema match carries no signal.
 */
const MIN_CONSOLIDATED_COLUMNS = 3;

export function columnIndexToLetter(col: number): string {
  let index = col + 1;
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

/** Ordered labels of a header run, lowercased for comparison. */
function schemaOf(run: HeaderRun): string[] {
  const labels: string[] = [];
  for (let col = run.startCol; col <= run.endCol; col += 1) {
    labels.push((run.labels.get(col) ?? '').trim().toLowerCase());
  }
  return labels;
}

function schemasMatch(a: string[], b: string[]): boolean {
  return a.length === b.length && a.every((label, i) => label === b[i]);
}

/**
 * A sheet name is safe to inline into a formula reference unquoted only when it
 * is a plain identifier. Anything else gets single quotes, with embedded single
 * quotes doubled — Excel's own escaping rule.
 */
export function formulaSheetRef(sheetName: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)) return sheetName;
  return `'${sheetName.replace(/'/g, "''")}'`;
}

/**
 * Build the spilling formula. One `LET`:
 *   `rows` — every source sheet's data range, each prefixed with its own name
 *   `keep` — per-row "does this row have anything in it?"
 *
 * `keep` deliberately tests the DATA columns (`DROP(rows,,1)`) and not the
 * origin label, which is present on every row by construction. It also tests
 * the whole row rather than the first column: the reported case had a user fill
 * Guest and Guest Name while leaving Unit No empty, and a first-column test
 * would have hidden that booking.
 */
export function buildConsolidationFormula(
  sourceSheets: string[],
  startCol: number,
  endCol: number,
): string {
  const firstLetter = columnIndexToLetter(startCol);
  const lastLetter = columnIndexToLetter(endCol);

  const blocks = sourceSheets.map((sheet) => {
    const ref = formulaSheetRef(sheet);
    const range = `${ref}!${firstLetter}2:${lastLetter}${SOURCE_ROW_LIMIT}`;
    const label = sheet.replace(/"/g, '""');
    return `HSTACK(EXPAND("${label}",ROWS(${range}),1,"${label}"),${range})`;
  });

  return (
    `=LET(rows,VSTACK(${blocks.join(',')}),` +
    `keep,BYROW(DROP(rows,,1),LAMBDA(r,COUNTA(r)>0)),` +
    `IFERROR(FILTER(rows,keep,""),""))`
  );
}

/**
 * What a host without dynamic arrays gets instead of the spilling view.
 *
 * Two per-cell fallbacks were built and both were wrong, which is worth
 * recording because the reasoning generalises:
 *
 *   1. 200 rows x `SMALL(IF(...))` array formula per cell = ~2,100 formulas.
 *      A live smoke test blew the response payload limit outright
 *      ("offset out of range ... received 17825794").
 *   2. 8 rows/sheet x positional INDEX = ~1,056 formulas. Payload survived,
 *      but a downstream stage still crashed on the volume.
 *
 * The deeper problem is that the dynamic-array form is ONE cell and *any*
 * per-cell equivalent is a thousand-plus — which bloats the change set, the
 * audit diff, and the Office.js apply, in exchange for eight rows per month.
 * That is a bad trade even when it works.
 *
 * So a legacy host gets an honest note instead. Crucially it loses only the
 * consolidated *detail grid*: the Monthly Totals rollup is plain SUMIF, so the
 * dashboard, the chart, and every per-month sheet still work exactly as on a
 * modern host. Telling the user why beats silently shipping either a `#NAME?`
 * or a thousand fragile formulas. TASKS.md #152.
 */
const LEGACY_NOTE =
  'Live consolidation of all months needs Excel 365 or 2021+. Your bookings stay on the month tabs, and the Monthly Summary above updates from them.';

/**
 * The compatibility path: one cell explaining why the live grid is absent.
 *
 * Deliberately a single SET_CELL. See `LEGACY_NOTE` for why a per-cell formula
 * grid was built twice, measured, and rejected.
 */
export function buildLegacyConsolidationNote(
  targetSheet: string,
  headerRow: number,
  col: number,
): SheetActionPayload[] {
  return [
    {
      type: 'SET_CELL',
      sheetName: targetSheet,
      row: headerRow + 1,
      col,
      value: LEGACY_NOTE,
    },
  ];
}

interface ConsolidationPlan {
  targetSheet: string;
  /** 0-based row the formula is written to — directly under the header. */
  row: number;
  col: number;
  sourceSheets: string[];
  formula: string;
  /** Source schema span, needed to build the compatible fallback. */
  sourceStartCol: number;
  sourceEndCol: number;
}

/**
 * Find a header row that consolidates other sheets: its first column names the
 * origin ("Month") and the rest is exactly the schema shared by two or more
 * other sheets in this batch.
 *
 * The exact-schema requirement is what keeps this from firing on Main's own
 * Monthly Totals table, whose header also starts with "Month" but continues
 * `Total Amount / Paid Amount / Pending Amount` — a rollup, not a consolidation.
 */
export function planConsolidation(actions: SheetActionPayload[]): ConsolidationPlan | null {
  const created = sheetsCreatedInBatch(actions);
  if (created.size === 0) return null;

  const runs = detectHeaderRuns(actions, created);
  if (runs.length === 0) return null;

  // Group candidate source sheets by their schema signature.
  const bySchema = new Map<string, { sheets: Set<string>; run: HeaderRun }>();
  for (const run of runs) {
    const key = schemaOf(run).join(SCHEMA_KEY_SEPARATOR);
    const entry = bySchema.get(key);
    if (entry) entry.sheets.add(run.sheetName);
    else bySchema.set(key, { sheets: new Set([run.sheetName]), run });
  }

  for (const run of runs) {
    const schema = schemaOf(run);
    // A consolidated table is recognised STRUCTURALLY, not by vocabulary: its
    // first column is an origin key, and its remaining columns are exactly the
    // schema shared by two or more sibling sheets this batch created.
    //
    // This used to additionally require that first label to be the literal
    // word "month", which silently limited the whole feature to calendar
    // workbooks — "a sheet per region / client / department / project", the
    // same shape with a different noun, got nothing. The exact-schema equality
    // below is what actually carries the signal (it is what stops Main's own
    // "Month | Total | Paid | Pending" rollup matching, since its columns are
    // NOT the month-sheet schema); the word never added safety, only
    // narrowness. TASKS.md #158.
    if (schema.length < MIN_CONSOLIDATED_COLUMNS) continue;
    if (!schema[0] || !schema[0].trim()) continue;

    const sourceSchema = schema.slice(1);
    const match = bySchema.get(sourceSchema.join(SCHEMA_KEY_SEPARATOR));
    if (!match) continue;

    const sourceSheets = [...match.sheets].filter(
      (name) => name.toLowerCase() !== run.sheetName.toLowerCase(),
    );
    if (sourceSheets.length < MIN_SOURCE_SHEETS) continue;

    // Order source sheets as the batch created them, so the consolidated view
    // follows the workbook's own tab order (Jan..Dec, North..West, Q1..Q4)
    // rather than alphabetically.
    const creationOrder = orderOfCreation(actions);
    sourceSheets.sort(
      (a, b) =>
        (creationOrder.get(a.toLowerCase()) ?? 0) - (creationOrder.get(b.toLowerCase()) ?? 0),
    );

    return {
      targetSheet: run.sheetName,
      row: run.row + 1,
      col: run.startCol,
      sourceSheets,
      formula: buildConsolidationFormula(
        sourceSheets,
        match.run.startCol,
        match.run.endCol,
      ),
      sourceStartCol: match.run.startCol,
      sourceEndCol: match.run.endCol,
    };
  }

  return null;
}

function orderOfCreation(actions: SheetActionPayload[]): Map<string, number> {
  const order = new Map<string, number>();
  let index = 0;
  for (const action of actions) {
    if (action.type !== 'ADD_SHEET' && action.type !== 'CREATE_SHEET') continue;
    const name = String(action.name ?? action.sheetName ?? '').trim().toLowerCase();
    if (name && !order.has(name)) order.set(name, index++);
  }
  return order;
}

/**
 * Append the consolidation formula when this batch builds a table that is
 * plainly meant to aggregate its sibling sheets. Append-only, one action.
 */
export function applyConsolidationPass(
  actions: SheetActionPayload[],
  options: { dynamicArrays?: boolean } = {},
): SheetActionPayload[] {
  const plan = planConsolidation(actions);
  if (!plan) return actions;

  // Only take the dynamic-array path when the host has been PROVEN to support
  // it. `dynamicArrays === undefined` means the probe never ran (older add-in
  // build, probe failed, non-taskpane caller) — and an unprobed host must get
  // the compatible form, because a wrong guess here is a silent `#NAME?` and an
  // empty consolidated table. TASKS.md #152.
  if (options.dynamicArrays !== true) {
    return [
      ...actions,
      ...buildLegacyConsolidationNote(plan.targetSheet, plan.row - 1, plan.col),
    ];
  }

  return [
    ...actions,
    {
      type: 'SET_FORMULA',
      sheetName: plan.targetSheet,
      row: plan.row,
      col: plan.col,
      formula: plan.formula,
    },
  ];
}
