import { Action } from '../../agents/types/agent.types';

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
}
