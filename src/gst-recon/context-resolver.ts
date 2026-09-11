import { ReconRunContext } from './types';

export interface MissingField {
  field: keyof ReconRunContext | 'client' | 'clientGstin' | 'taxPeriod';
  chatPrompt: string;
}

/**
 * Collect missing run-context fields. Do not guess silently.
 */
export function resolveMissingContext(
  partial: Partial<ReconRunContext> & {
    clientGstin?: string;
    taxPeriod?: string;
    client?: { id?: string; name?: string };
  },
  proposedGstinFromSheet?: string,
): MissingField[] {
  const missing: MissingField[] = [];

  if (!partial.clientGstin) {
    if (proposedGstinFromSheet) {
      missing.push({
        field: 'clientGstin',
        chatPrompt: `I found GSTIN ${proposedGstinFromSheet} in the register — should I use this as the client GSTIN for this run?`,
      });
    } else {
      missing.push({
        field: 'clientGstin',
        chatPrompt: `Which client GSTIN should I reconcile against? (This should be the client's own GSTIN — not your CA firm's.)`,
      });
    }
  }
  if (!partial.taxPeriod) {
    missing.push({
      field: 'taxPeriod',
      chatPrompt: `Which period should I reconcile — a specific month, or a quarter?`,
    });
  }
  if (!partial.client?.name) {
    missing.push({
      field: 'client',
      chatPrompt: `Which client is this for?`,
    });
  }
  return missing;
}
