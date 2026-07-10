// cellix_backend/src/excel-ai/utils/tiered-toon.util.ts

import { RouterPath } from '../types/router.types';

const toon = require('@toon-format-cjs/toon') as {
  encode: (input: unknown, options?: Record<string, unknown>) => string;
};

/**
 * Problem solved: the full TOON payload sends ALL sheet rows to the LLM.
 * For large sheets (1000 rows, 50 cols), this is 50,000 cells the Planner
 * never looks at for a simple write operation.
 *
 * Strategy: tiered payload based on routing decision.
 *
 * | Route  | Payload                              | Token reduction |
 * |--------|--------------------------------------|-----------------|
 * | write  | headers + 5 sample rows + metadata   | ~70–80%         |
 * | data   | full sheet (DataQueryService needs it) | 0%             |
 * | export | full sheet (row matching needs it)     | 0%             |
 * | ask    | headers + 10 rows (enough for explain) | ~60%           |
 * | shortcut | empty (no LLM call needed)           | 100%            |
 */
export interface TieredToonOptions {
  route: RouterPath;
  /** Raw WorkbookContext as received from the frontend */
  workbookContext: any;
  /** Already-compressed TOON string from frontend (used as-is for data/export) */
  rawToonPayload?: string;
}

export interface TieredToonResult {
  /** The TOON string to send to the LLM */
  promptContext: string;
  /** How many rows were included (for logging) */
  includedRows: number;
  /** Whether we used the tiered (reduced) payload */
  isTiered: boolean;
}

const WRITE_SAMPLE_ROWS = 5;
const ASK_SAMPLE_ROWS = 10;

export function buildTieredToon(options: TieredToonOptions): TieredToonResult {
  const { route, workbookContext, rawToonPayload } = options;

  // Shortcut: no TOON needed — action is deterministic
  if (route === 'shortcut') {
    return { promptContext: '', includedRows: 0, isTiered: true };
  }

  // Data / export: need full sheet for accurate aggregation and row matching
  if (route === 'data' || route === 'export') {
    return {
      promptContext: rawToonPayload ?? toon.encode(workbookContext),
      includedRows: getTotalRows(workbookContext),
      isTiered: false,
    };
  }

  // Write: Planner only needs headers + sample rows to understand structure
  if (route === 'write') {
    const sliced = sliceWorkbookContext(workbookContext, WRITE_SAMPLE_ROWS);
    return {
      promptContext: toon.encode(sliced),
      includedRows: WRITE_SAMPLE_ROWS,
      isTiered: true,
    };
  }

  // Ask: a bit more data for explaining patterns in the data
  if (route === 'ask') {
    const sliced = sliceWorkbookContext(workbookContext, ASK_SAMPLE_ROWS);
    return {
      promptContext: toon.encode(sliced),
      includedRows: ASK_SAMPLE_ROWS,
      isTiered: true,
    };
  }

  // Fallback: send full payload
  return {
    promptContext: rawToonPayload ?? toon.encode(workbookContext),
    includedRows: getTotalRows(workbookContext),
    isTiered: false,
  };
}

/**
 * Returns a copy of workbookContext with each sheet's rows
 * trimmed to `maxRows`, preserving the header row + first N data rows.
 * Also injects metadata so the Planner knows the actual sheet size.
 */
function sliceWorkbookContext(ctx: any, maxRows: number): any {
  if (!ctx?.sheets) return ctx;

  return {
    ...ctx,
    sheets: ctx.sheets.map((sheet: any) => {
      const rows: any[][] = sheet.rows ?? [];
      const slicedRows = rows.slice(0, maxRows + 1); // +1 for header row

      return {
        ...sheet,
        rows: slicedRows,
        // Inject actual size so Planner doesn't assume it sees everything
        _meta: {
          totalRows: rows.length,
          totalColumns: sheet.columns?.length ?? (rows[0]?.length ?? 0),
          rowsIncluded: slicedRows.length,
          note: `Only first ${slicedRows.length} rows shown. Full sheet has ${rows.length} rows.`,
        },
      };
    }),
  };
}

function getTotalRows(ctx: any): number {
  if (!ctx?.sheets) return 0;
  return ctx.sheets.reduce((sum: number, s: any) => sum + (s.rows?.length ?? 0), 0);
}
