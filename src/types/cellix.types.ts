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

export interface WorkbookContext {
  sheets: SheetSnapshot[];
  activeSheet: string;
  selectedRange?: string;
  selectedValues?: (string | number | null)[][];
  namedRanges?: NamedRangeInfo[];
  tables?: TableInfo[];
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
