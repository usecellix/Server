// cellix_backend/src/excel-ai/types/router.types.ts

/**
 * The output of the LLM Router — a single structured decision
 * that replaces all the regex-based intent/find/shortcut routing chains.
 */
export type RouterPath =
  | 'shortcut'   // layout command — no Planner, instant action
  | 'data'       // read-only aggregation/find — DataQueryService
  | 'export'     // find + copy rows to new sheet — FindExportService
  | 'write'      // data modification — full Orchestrator pipeline
  | 'ask';       // read-only explanation — streamWithOpenAi

export interface RouterDecision {
  /** Which handler path to use */
  route: RouterPath;

  /**
   * For route='shortcut' only: the exact SheetAction type to dispatch.
   * The expanded shortcut-router.util.ts will handle the rest.
   */
  action?: string;

  /**
   * 0.0–1.0. Below 0.60 on a destructive write → ask clarification.
   * Otherwise always proceed and state assumption in the answer.
   */
  confidence: number;

  /** One-sentence explanation of the routing decision (for debug logs) */
  reasoning: string;

  /**
   * When the router infers something ambiguous, this is the assumption
   * it made — injected into the Planner/answer so the user knows.
   * e.g. "I'll treat 'Amount' as the column to sum."
   */
  assumption?: string;

  /** Complexity tier for route='write' only — 0/1/2/3 */
  complexity?: 0 | 1 | 2 | 3;

  /** Catalog action hint for route='write' (e.g. FORMULA_GEN, SORT_OR_FILTER) */
  actionHint?: string;

  /** Whether complexity came from regex classifier or LLM semantic fallback */
  matchedBy?: 'regex' | 'llm-fallback';

  /**
   * True when write-intent-guard overrode a non-write route (data/ask/etc.)
   * because the message deterministically implies a sheet mutation.
   */
  overridden?: boolean;
}

/** What we send to the LLM Router */
export interface RouterInput {
  message: string;
  mode: 'ask' | 'action' | 'plan';
  /** Sheet headers only — lightweight, not full TOON */
  sheetHeaders: string[];
  /** Active sheet name */
  activeSheet: string;
  /** Last 2 user messages for follow-up resolution */
  recentHistory?: string[];
}
