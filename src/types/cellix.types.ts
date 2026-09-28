export type LLMTier = 'low' | 'medium' | 'high';

export interface ModelConfig {
  tier: LLMTier;
  model: string;
  maxTokens: number;
  costPer1kPrompt: number;
  costPer1kCompletion: number;
}

// Pricing and maxTokens must track the models actually configured in .env
// (OPENROUTER_MODEL_LOW/MEDIUM/HIGH). These feed both the HIGH-tier cost-cap
// downgrade in model-router.ts and every audit-log/dashboard cost figure —
// stale values here make real spend invisible, not just cosmetically wrong.
//
// Re-priced 2026-09-10 for the GLM model swap (.env's 2026-09-07 change) —
// this file was never updated when that swap happened, so every cost
// estimate and the COST_CAP_USD downgrade gate had been silently computing
// against gpt-5 pricing (~$1.25/$10 per 1M) while the app actually called
// GLM models (~18.7x cheaper blended). Pricing below confirmed directly
// against OpenRouter's live /api/v1/models listing on 2026-09-10:
//   - low:    z-ai/glm-5.3-flash   — $0.15 / $0.50 per 1M tokens
//   - medium: z-ai/glm-5.3         — $1.40 / $4.40 per 1M tokens (glm-5.2,
//             .env's previous MEDIUM value, does not exist on OpenRouter —
//             fixed alongside this repricing, see .env's own note)
//   - high:   z-ai/glm-latest      — alias resolving to z-ai/glm-5.3,
//             same real pricing as medium
// The HIGH→MEDIUM COST_CAP_USD downgrade in model-router.ts is now a genuine
// no-op again (medium and high price identically, both being glm-5.3) —
// unlike the pre-fix gpt-5 collapse, this is because glm-5.3 is honestly the
// right model for both tiers today, not because pricing was never updated.
export const MODEL_CONFIGS: Record<LLMTier, ModelConfig> = {
  low: {
    tier: 'low',
    model: 'z-ai/glm-5.3-flash',
    maxTokens: 8192,
    costPer1kPrompt: 0.00015,
    costPer1kCompletion: 0.0005,
  },
  medium: {
    tier: 'medium',
    model: 'z-ai/glm-5.3',
    maxTokens: 8192,
    costPer1kPrompt: 0.0014,
    costPer1kCompletion: 0.0044,
  },
  high: {
    tier: 'high',
    model: 'z-ai/glm-5.3',
    maxTokens: 8192,
    costPer1kPrompt: 0.0014,
    costPer1kCompletion: 0.0044,
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
  /**
   * Read from Office.js's real `worksheet.visibility` client-side (TASKS.md
   * #257) — optional so an older/minimal context that never populated it
   * reads as "unknown", not "definitely visible".
   */
  isHidden?: boolean;
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
