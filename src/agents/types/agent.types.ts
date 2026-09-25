import { SheetActionPayload } from '../../excel-ai/types/sheet-actions.types';
import { FormulaInsights, FormulaValidationIssue } from '../../formula/formula.types';
import { ConditionalFormatRuleInfo, SheetCompressionMeta } from '../../types/cellix.types';

export type Action = SheetActionPayload;

export interface RangeDataToolRequest {
  name: 'get_range_data';
  sheet: string;
  range: string;
}

export interface WorkbookContext {
  activeSheetName: string;
  sheets: SheetContext[];
  namedRanges: { name: string; formula: string }[];
  tables: string[];
  /**
   * Existing rules already on the live sheet, preserved with their full
   * identifying detail (unlike `tables` above, which is collapsed to names
   * only) — `id` is required to target one for `MODIFY_CONDITIONAL_FORMAT`
   * (TASKS.md #38). Optional — absent/undefined means "none known", same
   * convention as `namedRanges ?? []`/`tables ?? []` elsewhere in this file.
   */
  conditionalFormats?: ConditionalFormatRuleInfo[];
  selectedRange?: string;
  onDemandFetchEnabled?: boolean;
  fetchedRanges?: { sheet: string; range: string; rowCount: number }[];
  verifierFeedback?: string;
  verifierIssues?: VerifierIssue[];
  formulaValidationFeedback?: string;
  formulaValidationIssues?: FormulaValidationIssue[];
  /**
   * Spec 18 / Spec 21 — structured prior turn writes (charts + cell ranges)
   * for "along with the current" and overwrite-refinement recognition.
   */
  priorTurnActions?: Array<{
    actionType: string;
    sheetName: string;
    affectedRange?: string;
    targetColumn?: string;
    turnIndex?: number;
    sourceRange?: string;
    sourceSheetName?: string;
    destStartCell?: string;
    destSheet?: string;
    chartId?: string;
    chartType?: string;
    groupByColumn?: string;
  }>;
  priorTurnActionsSummary?: string;
}

/**
 * Bold/italic/fontColor/fillColor for one cell — column-level granularity in
 * practice (broadcast from `ColumnMeta.format`, itself read from a column's
 * first data row, TASKS.md #64), not a genuine per-cell snapshot. Mirrors
 * `numberFormats`' own existing column-broadcast precedent.
 */
export interface CellFormatCell {
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  fillColor?: string;
}

export interface SheetContext {
  name: string;
  usedRange: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: string[][];
  numberFormats: string[][];
  /** Absent when the add-in build predates TASKS.md #64 — treat as "no data", not "no formatting". */
  formats?: CellFormatCell[][];
  structure: 'financial_model' | 'data_table' | 'report' | 'unknown';
  /**
   * 0-based index into `values` where column headers live. Not always 0 — sheets
   * with a title row above the table (common in exported reports) have headers
   * further down. The Executor must key off this instead of assuming row 0.
   */
  headerRowIndex: number;
  formulaInsights?: FormulaInsights;
  compressionMeta?: SheetCompressionMeta;
  dataTruncated?: boolean;
}

export interface SubTask {
  id: string;
  description: string;
  targetSheet: string;
  dependsOn: string[];
  estimatedActions: number;
  /** Optional nudge toward a native action type (e.g. COPY_FILTERED_RANGE). */
  suggestedActionType?: string;
  /**
   * Column headers the USER spelled out for this subtask's sheet, verbatim and
   * in order — set by `applyBuildSpecToSubtasks` (Phase 1 of
   * LONG_PROMPT_RELIABILITY_PLAN.md) and enforced by SpecConformanceChecker.
   * Persists with the run because `subtasks` is stored as-is.
   */
  expectedHeaders?: string[];
  /**
   * Set by `splitSpecPinnedSubtasks` — this subtask's actions (create sheet +
   * header row + table) are built entirely by code from `expectedHeaders`,
   * with no Executor/LLM call at all. `agenticLoop.service.ts` short-circuits
   * on this flag; `ComputedColumnChecker` exempts it (its formula, if any,
   * belongs to the dependent subtask this one was split from).
   */
  isDeterministicHeaderStep?: boolean;
  /**
   * The FULL header row this step must build — the user's columns plus any
   * computed ones the planner interleaved — resolved ONCE at split time from
   * the original subtask's description. TASKS.md #292.
   *
   * Carried explicitly because the split gives the header step a short,
   * synthesized description of its own; re-deriving the row from THAT at
   * execution time silently falls back to the narrower `expectedHeaders`, so
   * the table gets built narrower than the rest step was told it would be.
   */
  resolvedHeaderRow?: string[];
  /**
   * The complete, code-built actions for this step — no Executor/LLM call.
   * Set at plan time (e.g. the ledger dashboard, `dashboard-builder.util.ts`,
   * TASKS.md #327) and persisted with the run like the rest of the subtask;
   * `agenticLoop.service.ts` applies them verbatim.
   */
  deterministicActions?: Action[];
}

/** A step whose actions are built by code rather than the Executor. */
export function isDeterministicStep(subtask: SubTask): boolean {
  return Boolean(subtask.isDeterministicHeaderStep || subtask.deterministicActions?.length);
}

export interface PlannerOutput {
  subtasks: SubTask[];
  clarificationsNeeded: string[];
  confidence: 'high' | 'medium' | 'low';
  reasoning: string;
}

/**
 * Two-pass planning (TASKS.md #191) — a coarse, cheap-to-generate top-level
 * item the Planner's first pass identifies, expanded into real `SubTask[]` by
 * a SEPARATE, phase-scoped second call. Exists because a single-pass plan for
 * a large compound build (12 month sheets + a multi-section dashboard) has to
 * describe everything in one JSON response, and repeatedly truncates even
 * after exhausting the last-resort token ceiling (TASKS.md #170/#187/#190) —
 * splitting the DESCRIBING into small, independent calls is the actual fix,
 * not a bigger ceiling that the next larger request outgrows again.
 */
export interface PlanPhase {
  id: string;
  /** Free-text description of what this phase covers — fed to the expansion
   *  call, not the Executor. Not a SubTask.description. */
  kind: string;
  targetSheet: string;
  /**
   * When set, this phase covers ONE repeated structure applied once per
   * entry (e.g. ["January", ..., "December"] for 12 near-identical month
   * sheets) — the expansion call is told to produce one group of subtasks
   * PER entry, not a single subtask describing all of them at once.
   */
  repeatFor?: string[];
  /** Ids of other PHASES this one depends on — resolved to real subtask ids
   *  once both phases have been expanded (see `stitchPhaseSubtasks`). */
  dependsOn: string[];
}

export interface CoarsePlanOutput {
  phases: PlanPhase[];
  clarificationsNeeded: string[];
  confidence: 'high' | 'medium' | 'low';
  reasoning: string;
}

/**
 * An action the Executor emitted that could not be normalized into a usable action.
 * Carried out of normalization so it is logged and verified against — never silently discarded.
 */
export interface DroppedAction {
  /** The `type` the model emitted, when it was a string at all. */
  rawType: string | null;
  reason: 'not-an-object' | 'unknown-type' | 'missing-required-fields';
}

export interface ExecutorOutput {
  subtaskId: string;
  actions: Action[];
  isDone: boolean;
  nextStep?: string;
  toolRequest?: RangeDataToolRequest;
  /** False when the executor needed a JSON parse retry. */
  parsedOnFirstAttempt?: boolean;
  /** Actions the model emitted that normalization could not use. Empty/absent when all were kept. */
  droppedActions?: DroppedAction[];
}

export interface VerifierIssue {
  severity: 'error' | 'warning';
  actionIndex?: number;
  subtaskId?: string;
  description: string;
  suggestion: string;
}

export interface VerifierSubtaskResult {
  subtaskId: string;
  passed: boolean;
  feedback: string;
  issues: VerifierIssue[];
  /**
   * True when the verifier response was truncated before this subtask —
   * needs re-verification only, not re-execution of a prior pass.
   */
  inconclusive?: boolean;
}

export interface VerifierOutput {
  passed: boolean;
  feedback: string;
  issues: VerifierIssue[];
  subtaskResults: VerifierSubtaskResult[];
  /** @deprecated Verifier must not invent actions — executor retries with feedback instead. */
  revisedActions?: Action[];
}

export interface AgentRunOptions {
  prompt: string;
  context: WorkbookContext;
  conversationHistory?: { role: 'user' | 'assistant'; content: string }[];
  maxRetries?: number;
  promptContext?: string;
  conversationId?: string;
  correlationId?: string;
  toolEmit?: (event: string, data: Record<string, unknown>) => void;
  routerAssumption?: string;
  /** Router complexity tier (0–3) — keys Planner max_tokens budget. */
  complexity?: 0 | 1 | 2 | 3;
  /**
   * Called as each execution wave completes, so the caller can emit an Accept
   * card for finished work instead of waiting for the whole run. TASKS.md #174.
   */
  onWaveComplete?: (waveActions: Action[], waveIndex: number) => Promise<void>;
  /**
   * A plan the caller has ALREADY computed, handed over instead of being
   * re-derived. Set by the stepwise gate when it plans, finds a single wave and
   * declines — without this the one-shot path pays for a second full Planner
   * call on every simple request. Must be a copy the caller no longer holds a
   * reference to. TASKS.md #196.
   */
  precomputedPlan?: PlannerOutput;
  /**
   * Fires when the client disconnects or aborts (e.g. the "Stop" button) —
   * checked between execution waves so a cancelled run stops burning LLM
   * calls/time instead of running to completion against a response nobody
   * is reading anymore.
   */
  abortSignal?: AbortSignal;
}
