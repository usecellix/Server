import { Action } from '../../agents/types/agent.types';
import {
  DomainException,
  SourceRef,
} from '../../domain-tools/types/domain-tool.types';

export type { DomainException, SourceRef };

export type ChangeSetStatus = 'previewed' | 'applied' | 'reverted';

export interface CellSnapshot {
  value: unknown;
  formula: string;
  format: string;
}

export interface CellChange {
  cell: string;
  sheet: string;
  before: unknown;
  after: unknown;
  formula?: string;
  isHardcoded: boolean;
  /** Optional citations — undefined for Tier 0/1 is expected */
  sourceRefs?: SourceRef[];
  /** Domain-tool exception flags — rendered distinctly in the UI */
  exceptionFlags?: DomainException[];
}

export interface ChangeSetRecord {
  changeSetId: string;
  conversationId: string;
  traceId: string;
  timestamp: Date;
  prompt: string;
  beforeState: Record<string, CellSnapshot>;
  changes: CellChange[];
  actions: Action[];
  status: ChangeSetStatus;
  appliedAt?: Date;
  revertedAt?: Date;
  /** Aggregate provenance for the change set (domain-tool confidence, etc.) */
  provenanceConfidence?: number;
}
