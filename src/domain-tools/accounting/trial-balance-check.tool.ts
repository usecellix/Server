import { DomainTool, SourceRef } from '../types/domain-tool.types';

export interface TrialBalanceRow {
  accountCode: string;
  accountName: string;
  debit: number;
  credit: number;
  sourceRowRef: SourceRef;
}

export interface TrialBalanceCheckInput {
  rows: TrialBalanceRow[];
  tolerance: number;
}

export interface TrialBalanceCheckOutput {
  totalDebit: number;
  totalCredit: number;
  difference: number;
  balanced: boolean;
}

/**
 * STUB — trial balance check is pure arithmetic; still gated until fixture sign-off.
 */
export const trialBalanceCheck: DomainTool<
  TrialBalanceCheckInput,
  TrialBalanceCheckOutput
> = (_input) => {
  throw new Error(
    'Not implemented — requires CA-reviewed trial balance check spec before production use.',
  );
};
