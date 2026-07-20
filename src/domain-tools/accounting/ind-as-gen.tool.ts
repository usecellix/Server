import { DomainTool, SourceRef } from '../types/domain-tool.types';

export interface IndAsGenInput {
  trialBalanceRows: Array<{
    accountCode: string;
    accountName: string;
    debit: number;
    credit: number;
    sourceRowRef: SourceRef;
  }>;
  standard: 'IndAS';
}

export interface IndAsLine {
  accountCode: string;
  accountName: string;
  presentationLine: string;
  amount: number;
  sourceRowRef: SourceRef;
}

export interface IndAsGenOutput {
  lines: IndAsLine[];
}

/**
 * STUB — Ind-AS presentation mapping requires CA-reviewed rules before production use.
 */
export const indAsGen: DomainTool<IndAsGenInput, IndAsGenOutput> = (_input) => {
  throw new Error(
    'Not implemented — requires CA-reviewed Ind-AS mapping spec before production use.',
  );
};
