import { DomainTool, SourceRef } from '../types/domain-tool.types';

export interface CostAllocationInput {
  costPool: number;
  bases: Array<{
    costCenter: string;
    allocationBase: number;
    sourceRowRef: SourceRef;
  }>;
}

export interface CostAllocationLine {
  costCenter: string;
  allocatedAmount: number;
  share: number;
  sourceRowRef: SourceRef;
}

export interface CostAllocationOutput {
  lines: CostAllocationLine[];
  totalAllocated: number;
}

/**
 * STUB — cost allocation arithmetic must remain deterministic (never LLM-computed).
 */
export const costAllocation: DomainTool<CostAllocationInput, CostAllocationOutput> = (
  _input,
) => {
  throw new Error(
    'Not implemented — requires CA-reviewed cost allocation spec before production use.',
  );
};
