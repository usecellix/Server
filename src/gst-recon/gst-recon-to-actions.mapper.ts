import { SheetActionPayload } from '../excel-ai/types/sheet-actions.types';
import { SourceRef } from '../domain-tools/types/domain-tool.types';
import { GstResultRow } from '../domain-tools/gst/match-passes';
import { ItcComputeOutput } from '../domain-tools/gst/itc-compute.tool';
import { Gstr3bVs2bOutput } from '../domain-tools/gst/gstr3b-vs-2b.tool';

export interface ReconSummaryBlock {
  total_pr_rows: number;
  total_portal_rows: number;
  total_ims_rows: number;
  exact_matched: number;
  partial_matched: number;
  credit_notes: number;
  pr_only: number;
  portal_only: number;
  ims_rejected: number;
  ims_pending: number;
  ims_auto_accept: number;
  rcm_flagged: number;
  itc_matched: number;
  itc_at_risk: number;
  rcm_payable: number;
  itc_ims_rejected: number;
  itc_ims_pending: number;
}

export interface ReconRowDto {
  pr_ref: string | null;
  portal_ref: string | null;
  status: string;
  pass: number | null;
  confidence: number;
  difference: string | null;
  diff_type: string | null;
  itc_amount: number;
  ims_status: string | null;
  rcm_flag: boolean;
  invoice_number: string | null;
  gstin: string | null;
}

function refLabel(ref?: SourceRef): string | null {
  if (!ref) return null;
  return `${ref.documentId}:row_${ref.rowOrLine}`;
}

export function buildSummary(
  resultRows: GstResultRow[],
  itc: ItcComputeOutput,
  prCount: number,
  portalCount: number,
  imsCount: number,
): ReconSummaryBlock {
  const count = (status: string) => resultRows.filter((r) => r.status === status).length;
  return {
    total_pr_rows: prCount,
    total_portal_rows: portalCount,
    total_ims_rows: imsCount,
    exact_matched: count('MATCHED'),
    partial_matched: count('PARTIAL') + count('IMS_AUTO_ACCEPT'),
    credit_notes: count('CREDIT_NOTE'),
    pr_only: count('PR_ONLY'),
    portal_only: count('PORTAL_ONLY') + count('IMS_ONLY'),
    ims_rejected: count('IMS_REJECTED'),
    ims_pending: count('IMS_PENDING'),
    ims_auto_accept: resultRows.filter(
      (r) => r.imsStatus === 'AutoAccepted' || r.status === 'IMS_AUTO_ACCEPT',
    ).length,
    rcm_flagged: count('RCM'),
    itc_matched: itc.totalItcClaimable,
    itc_at_risk: itc.totalItcAtRisk,
    rcm_payable: itc.totalRcmPayable,
    itc_ims_rejected: itc.totalImsRejected,
    itc_ims_pending: itc.totalImsPending,
  };
}

export function mapResultRows(resultRows: GstResultRow[]): ReconRowDto[] {
  return resultRows.map((r) => {
    const row = r.registerRow ?? r.portalRow ?? r.imsRow;
    return {
      pr_ref: refLabel(r.registerRow?.sourceRowRef),
      portal_ref: refLabel(r.portalRow?.sourceRowRef ?? r.imsRow?.sourceRowRef),
      status: r.status,
      pass: r.pass,
      confidence: r.confidence,
      difference: r.difference ?? null,
      diff_type: r.diffType ?? null,
      itc_amount: r.itcAmount,
      ims_status: r.imsStatus ?? null,
      rcm_flag: r.rcmFlag,
      invoice_number: row?.invoiceNumber ?? null,
      gstin: row?.gstin ?? null,
    };
  });
}

export function mapToSheetActions(params: {
  sheetName: string;
  period?: string;
  gstin?: string;
  reconType: string;
  summary: ReconSummaryBlock;
  rows: ReconRowDto[];
  runAt: string;
  relativeTo?: string;
}): SheetActionPayload[] {
  const { sheetName, period, gstin, reconType, summary, rows, runAt, relativeTo } = params;

  const headers = [
    'Status',
    'Pass',
    'GSTIN',
    'Invoice No',
    'ITC Amount',
    'Confidence',
    'Difference',
    'Diff Type',
    'IMS Status',
    'RCM',
    'PR Ref',
    'Portal Ref',
  ];

  const tableRows: unknown[][] = rows.map((r) => [
    r.status,
    r.pass ?? '',
    r.gstin ?? '',
    r.invoice_number ?? '',
    r.itc_amount,
    r.confidence,
    r.difference ?? '',
    r.diff_type ?? '',
    r.ims_status ?? '',
    r.rcm_flag ? 'Y' : '',
    r.pr_ref ?? '',
    r.portal_ref ?? '',
  ]);

  // Summary block rows first (as part of table for ICAI working paper style)
  const summaryBlock: unknown[][] = [
    ['CELLIX GST Reconciliation Report', '', '', '', '', '', '', '', '', '', '', ''],
    [`Type: ${reconType}`, `Period: ${period ?? ''}`, `GSTIN: ${gstin ?? ''}`, `Run: ${runAt}`, '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', ''],
    ['SUMMARY', '', '', '', '', '', '', '', '', '', '', ''],
    ['Purchase Register Rows', summary.total_pr_rows, '', '', '', '', '', '', '', '', '', ''],
    ['Portal / GSTR Rows', summary.total_portal_rows, '', '', '', '', '', '', '', '', '', ''],
    ['IMS Entries', summary.total_ims_rows, '', '', '', '', '', '', '', '', '', ''],
    ['Exact Matched', summary.exact_matched, 'ITC matched', summary.itc_matched, '', '', '', '', '', '', '', ''],
    ['Partial Matched', summary.partial_matched, '', '', '', '', '', '', '', '', '', ''],
    ['Credit Notes', summary.credit_notes, '', '', '', '', '', '', '', '', '', ''],
    ['PR Only', summary.pr_only, 'ITC at risk', summary.itc_at_risk, '', '', '', '', '', '', '', ''],
    ['Portal Only', summary.portal_only, '', '', '', '', '', '', '', '', '', ''],
    ['IMS Rejected', summary.ims_rejected, 'ITC lost', summary.itc_ims_rejected, '', '', '', '', '', '', '', ''],
    ['IMS Pending', summary.ims_pending, 'ITC deferred', summary.itc_ims_pending, '', '', '', '', '', '', '', ''],
    ['RCM Flags', summary.rcm_flagged, 'RCM payable', summary.rcm_payable, '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', ''],
    ['DETAIL', '', '', '', '', '', '', '', '', '', '', ''],
    headers,
    ...tableRows,
    ['', '', '', '', '', '', '', '', '', '', '', ''],
    ['CELLIX | Reconciliation aid. CA to verify.', '', '', '', '', '', '', '', '', '', '', ''],
  ];

  // WRITE_TABLE uses first row as headers — put markers as fake header then content rows
  const writeHeaders = summaryBlock[0].map((c, i) => (i === 0 ? String(c) : `C${i + 1}`));
  const writeRows = summaryBlock.slice(1);

  const actions: SheetActionPayload[] = [
    {
      type: 'CREATE_SHEET',
      sheetName,
      name: sheetName,
      relativeTo,
      position: 'after',
    },
    {
      type: 'WRITE_TABLE',
      sheetName,
      headers: writeHeaders as string[],
      rows: writeRows,
    },
  ];

  return actions;
}

export function map3bVs2bToActions(params: {
  sheetName: string;
  period?: string;
  gstin?: string;
  result: Gstr3bVs2bOutput;
  runAt: string;
  relativeTo?: string;
}): SheetActionPayload[] {
  const { sheetName, period, gstin, result, runAt, relativeTo } = params;
  const headers = ['Metric', 'IGST', 'CGST', 'SGST', 'Total', 'Notes'];
  const rows: unknown[][] = [
    [
      'Variance (3B − 2B)',
      result.variances.igst,
      result.variances.cgst,
      result.variances.sgst,
      result.variances.total,
      result.status,
    ],
    ['Excess claim', '', '', '', result.excessClaim, result.drc01cRisk ? 'DRC-01C RISK' : ''],
    ['Short claim', '', '', '', result.shortClaim, ''],
    ...result.messages.map((m) => ['Note', '', '', '', '', m]),
    ['', '', '', '', '', 'CELLIX | Reconciliation aid. CA to verify.'],
  ];

  return [
    {
      type: 'CREATE_SHEET',
      sheetName,
      name: sheetName,
      relativeTo,
      position: 'after',
    },
    {
      type: 'WRITE_TABLE',
      sheetName,
      headers,
      rows: [
        [`Period: ${period ?? ''} | GSTIN: ${gstin ?? ''} | Run: ${runAt}`, '', '', '', '', ''],
        ...rows,
      ],
    },
  ];
}
