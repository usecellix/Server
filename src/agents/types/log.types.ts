export interface AgentLogEvent {
  correlationId: string;
  agent: 'planner' | 'executor' | 'verifier' | 'workbook';
  model: string;
  durationMs: number;
  success: boolean;
  tokenUsage?: number;
  rawResponse?: string;
  parsedResponse?: unknown;
  error?: string;
}
