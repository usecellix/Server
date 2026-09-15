import { ReconType } from './types';

export type GstReconIntent = 'GST_RECON_PURCHASE' | 'GST_RECON_SALES';

export interface GstReconIntentPayload {
  intent: GstReconIntent;
  extractedClientName?: string;
  extractedGstin?: string;
  extractedPeriod?: string;
  extractedFinancialYear?: string;
}

export type ReconOrchestratorResult =
  | { kind: 'needs_input'; prompts: string[] }
  | { kind: 'error'; message: string }
  | { kind: 'chat_reply'; message: string }
  | {
      kind: 'needs_sheet_data';
      booksSheet: string;
      portalSheet: string;
      reconType: ReconType;
      clientGstin: string;
      period: string;
      clientName?: string;
      financialYear?: string;
    }
  | {
      kind: 'action_payload';
      summary: unknown;
      matchResults: unknown;
      actions: unknown;
      suggestedSheetName: string;
      jobId: string;
      clientGstin: string;
      reconType: ReconType;
      crossGstinExceptionCount: number;
      auditLogId: string | null;
    };
