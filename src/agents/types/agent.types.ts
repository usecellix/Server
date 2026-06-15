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
  toolEmit?: (event: string, data: Record<string, unknown>) => void;
}
