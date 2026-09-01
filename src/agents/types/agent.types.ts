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
}

export interface PlannerOutput {
  subtasks: SubTask[];
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
}
