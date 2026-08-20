export type LLMTier = 'low' | 'medium' | 'high';

export interface ModelConfig {
  tier: LLMTier;
  model: string;
  maxTokens: number;
  costPer1kPrompt: number;
  costPer1kCompletion: number;
}

export const MODEL_CONFIGS: Record<LLMTier, ModelConfig> = {
  low: {
    tier: 'low',
    model: 'google/gemini-flash-1.5',
    maxTokens: 2048,
    costPer1kPrompt: 0.000075,
    costPer1kCompletion: 0.0003,
  },
  medium: {
    tier: 'medium',
    model: 'openai/gpt-4o-mini',
    maxTokens: 4096,
    costPer1kPrompt: 0.00015,
    costPer1kCompletion: 0.0006,
  },
  high: {
    tier: 'high',
    model: 'openai/gpt-4o',
    maxTokens: 8192,
    costPer1kPrompt: 0.0025,
    costPer1kCompletion: 0.01,
  },
};

export type ColumnDetectedType =
  | 'date'
  | 'currency'
  | 'number'
  | 'text'
  | 'boolean'
  | 'unknown';

export interface ColumnMeta {
  index: number;
  header?: string;
  sampleValues: (string | number | null)[];
  detectedType?: ColumnDetectedType | string;
  numberFormat?: string;
  /**
   * Bold/italic/fontColor/fillColor read from the column's first data row
   * (TASKS.md #64) — column-level granularity, same as `numberFormat` above,
   * not a genuine per-cell snapshot. Feeds revert's format restoration.
   */
  format?: {
    bold?: boolean;
    italic?: boolean;
    fontColor?: string;
    fillColor?: string;
  };
}

export interface ConversationTurn {
  role: 'user' | 'assistant' | string;
  content: string;
}

export type SheetStructure = 'financial_model' | 'data_table' | 'report' | 'unknown';

export interface SheetCompressionMeta {
  originalRowCount: number;
  compressedRowCount: number;
  truncated: boolean;
  onDemandFetchEnabled: boolean;
  includedRowIndices?: number[];
}

export interface SheetSnapshot {
  sheetName: string;
  usedRange: string;
  rowCount: number;
  colCount: number;
  headers: string[];
  /**
   * 0-based row index within sampleData where `headers` actually lives, when the
   * add-in detected it. Undefined on older add-in builds — callers should treat
   * that as "unknown" and re-detect from sampleData rather than assuming row 0.
   */
  headerRowIndex?: number;
  sampleData: (string | number | null)[][];
  columnMeta?: ColumnMeta[];
  structure?: SheetStructure;
  formulaSummary?: string;
  compressionMeta?: SheetCompressionMeta;
}

export interface NamedRangeInfo {
  name: string;
  formula: string;
  type?: string;
}

export interface TableInfo {
  name: string;
  sheetName: string;
  range?: string;
  hasHeaders?: boolean;
  columnNames: string[];
}

/**
 * A conditional-format rule already present on the live sheet, read back via
 * Office.js (`client/src/context/workbookReader.ts`, TASKS.md #38) — not
 * limited to rules Cellix itself applied. `id` is what lets a follow-up
 * request target this specific rule (`MODIFY_CONDITIONAL_FORMAT`) instead of
 * stacking a duplicate `CONDITIONAL_FORMAT` on top.
 */
export interface ConditionalFormatRuleInfo {
  id: string;
  sheetName: string;
  range: string;
  ruleKind: 'cellValue' | 'formula' | 'topBottom' | 'colorScale' | 'other';
  summary: string;
}

export interface WorkbookContext {
  sheets: SheetSnapshot[];
  activeSheet: string;
  selectedRange?: string;
  selectedValues?: (string | number | null)[][];
  namedRanges?: NamedRangeInfo[];
  tables?: TableInfo[];
  conditionalFormats?: ConditionalFormatRuleInfo[];
  prompt_context?: string;
}

export interface ClarificationPayload {
  question: string;
  suggestions?: string[];
  ambiguityScore: number;
}

export type UserIntent =
  | 'create_data'
  | 'modify_data'
  | 'format'
  | 'formula'
  | 'sort_filter'
  | 'analyze'
  | 'delete'
  | 'other';

export interface AuditLogEntry {
  id: string;
  traceId: string;
  timestamp: string;
  model: string;
  tier: LLMTier;
  intent: string;
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  estimatedCostUsd: number;
  latencyMs: number;
  success: boolean;
  errorCode?: string;
  actionsCount?: number;
}
