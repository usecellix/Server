import { SheetActionPayload } from '../../excel-ai/types/sheet-actions.types';
import { FormulaInsights, FormulaValidationIssue } from '../../formula/formula.types';
import { SheetCompressionMeta } from '../../types/cellix.types';

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
  selectedRange?: string;
  onDemandFetchEnabled?: boolean;
  fetchedRanges?: { sheet: string; range: string; rowCount: number }[];
  verifierFeedback?: string;
  verifierIssues?: VerifierIssue[];
  formulaValidationFeedback?: string;
  formulaValidationIssues?: FormulaValidationIssue[];
  /**
   * Spec 18 — structured prior CREATE_CHART / AGGREGATE_TABLE records for
   * follow-ups like "along with the current".
   */
  priorTurnActions?: Array<{
    actionType: string;
    sheetName: string;
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

export interface SheetContext {
  name: string;
  usedRange: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: string[][];
  numberFormats: string[][];
  structure: 'financial_model' | 'data_table' | 'report' | 'unknown';
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

export interface ExecutorOutput {
  subtaskId: string;
  actions: Action[];
  isDone: boolean;
  nextStep?: string;
  toolRequest?: RangeDataToolRequest;
  /** False when the executor needed a JSON parse retry. */
  parsedOnFirstAttempt?: boolean;
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
