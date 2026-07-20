import { DomainTool } from './types/domain-tool.types';
import { gstMatch } from './gst/gst-match.tool';
import { itcCompute } from './gst/itc-compute.tool';
import { tds26asMatch } from './tds/tds-26as-match.tool';
import { bankRecon } from './reconciliation/bank-recon.tool';
import { indAsGen } from './accounting/ind-as-gen.tool';
import { trialBalanceCheck } from './accounting/trial-balance-check.tool';
import { costAllocation } from './accounting/cost-allocation.tool';

/**
 * Tools ExecutorAgent may call for domain-flagged subtasks.
 * Implementations are deterministic TypeScript — never LLM free-text arithmetic.
 */
export const domainToolRegistry: Record<string, DomainTool<unknown, unknown>> = {
  gst_match: gstMatch as DomainTool<unknown, unknown>,
  itc_compute: itcCompute as DomainTool<unknown, unknown>,
  tds_26as_match: tds26asMatch as DomainTool<unknown, unknown>,
  bank_recon: bankRecon as DomainTool<unknown, unknown>,
  ind_as_gen: indAsGen as DomainTool<unknown, unknown>,
  trial_balance_check: trialBalanceCheck as DomainTool<unknown, unknown>,
  cost_allocation: costAllocation as DomainTool<unknown, unknown>,
};

export const DOMAIN_TOOL_NAMES = Object.keys(domainToolRegistry) as Array<
  keyof typeof domainToolRegistry
>;

export const DOMAIN_TOOL_REGISTRY = Symbol('DOMAIN_TOOL_REGISTRY');
