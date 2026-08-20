import { Action } from '../../agents/types/agent.types';
import { CheckpointStatus, CheckpointTrigger } from '../schemas/checkpoint.schema';

export type { CheckpointStatus, CheckpointTrigger };

export interface CheckpointRecord {
  checkpointId: string;
  workbookId: string;
  conversationId: string;
  label: string;
  trigger: CheckpointTrigger;
  anchorChangeSetId: string;
  createdAt: Date;
  status: CheckpointStatus;
  restoredAt?: Date;
}

export interface RestoreResult {
  checkpoint: CheckpointRecord;
  /** change_sets docs marked 'reverted' by this restore, newest-first. */
  revertedChangeSetIds: string[];
  /**
   * The concatenated inverse actions, in application order, for the frontend to
   * apply via Office.js (AD-1: the backend never writes to the live workbook —
   * this mirrors ChangeSetService.revert's own return shape exactly).
   */
  inverseActions: Action[];
}
