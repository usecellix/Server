import {
  DEFAULT_GST_MATCH_SETTINGS,
  DomainTool,
  DomainToolResult,
  GstMatchSettings,
  MatchedPair,
  NormalizedInvoiceRow,
  SourceRef,
} from '../types/domain-tool.types';
import {
  createWorkingSet,
  finalizeUnmatched,
  GstResultRow,
  runPassCdn,
  runPassExact,
  runPassFuzzy,
  runPassIms,
  runPassRcm,
  toMatchedPairs,
} from './match-passes';

export interface GstMatchInput {
  purchaseRegister: NormalizedInvoiceRow[];
  /** GSTR-2B or GSTR-2A rows */
  gstr2b: NormalizedInvoiceRow[];
  /** Optional IMS rows for pass 5 */
  imsRows?: NormalizedInvoiceRow[];
  /** @deprecated prefer settings */
  matchKeys?: Array<'gstin' | 'invoiceNumber' | 'invoiceDate'>;
  /** @deprecated prefer settings.amountToleranceAbs */
  amountTolerance?: number;
  settings?: Partial<GstMatchSettings>;
}

export interface GstMatchOutput {
  matched: MatchedPair[];
  partialMatch: MatchedPair[];
  missingIn2B: NormalizedInvoiceRow[];
  missingInRegister: NormalizedInvoiceRow[];
  /** Full categorized result rows for recon report */
  resultRows: GstResultRow[];
  settings: GstMatchSettings;
}

function collectSourceRefs(rows: GstResultRow[]): SourceRef[] {
  const refs: SourceRef[] = [];
  for (const r of rows) {
    if (r.registerRow) refs.push(r.registerRow.sourceRowRef);
    if (r.portalRow) refs.push(r.portalRow.sourceRowRef);
    if (r.imsRow) refs.push(r.imsRow.sourceRowRef);
  }
  return refs;
}

/**
 * 6-pass pipeline P1–P5 (deterministic). Pass 6 LLM is orchestrator-only.
 */
export const gstMatch: DomainTool<GstMatchInput, GstMatchOutput> = (input) => {
  const settings: GstMatchSettings = {
    ...DEFAULT_GST_MATCH_SETTINGS,
    ...(input.amountTolerance != null
      ? { amountToleranceAbs: input.amountTolerance }
      : {}),
    ...input.settings,
  };

  const ws = createWorkingSet(input.purchaseRegister ?? [], input.gstr2b ?? []);

  runPassExact(ws);
  runPassFuzzy(ws, settings);
  runPassCdn(ws);
  runPassRcm(ws, settings);
  runPassIms(ws, input.imsRows ?? [], settings);
  finalizeUnmatched(ws);

  const matched = toMatchedPairs(ws.results.filter((r) => r.status === 'MATCHED'));
  const partialMatch = toMatchedPairs(
    ws.results.filter(
      (r) =>
        r.status === 'PARTIAL' ||
        r.status === 'CREDIT_NOTE' ||
        r.status === 'IMS_PENDING' ||
        r.status === 'IMS_AUTO_ACCEPT',
    ),
  );
  const missingIn2B = ws.results
    .filter((r) => r.status === 'PR_ONLY' || r.status === 'RCM' || r.status === 'IMS_REJECTED')
    .map((r) => r.registerRow!)
    .filter(Boolean);
  const missingInRegister = ws.results
    .filter((r) => r.status === 'PORTAL_ONLY' || r.status === 'IMS_ONLY')
    .map((r) => r.portalRow ?? r.imsRow!)
    .filter(Boolean);

  const exactCount = ws.results.filter((r) => r.status === 'MATCHED').length;
  const total = Math.max(ws.results.length, 1);
  const confidence = Math.min(1, 0.5 + (exactCount / total) * 0.5);

  const result: DomainToolResult<GstMatchOutput> = {
    data: {
      matched,
      partialMatch,
      missingIn2B,
      missingInRegister,
      resultRows: ws.results,
      settings,
    },
    confidence,
    exceptions: ws.exceptions,
    sourceRefs: collectSourceRefs(ws.results),
  };
  return result;
};
