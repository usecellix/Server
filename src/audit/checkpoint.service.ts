import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { randomUUID } from 'crypto';
import { Model } from 'mongoose';
import { Action } from '../agents/types/agent.types';
import { isDestructiveActionType } from '../agents/verifier-skip.policy';
import { WorkflowTraceService } from '../common/logging/workflow-trace.service';
import {
  beforeStateToInverseActions,
  diffShadowsFully,
  shadowFromBeforeState,
  structuralOpsToInverseActions,
} from './diff.engine';
import { RestoreVerificationError } from './errors/restore-verification.error';
import { virtualApply } from '../virtual/virtualApply';
import { Checkpoint, CheckpointDocument } from './schemas/checkpoint.schema';
import { ChangeSet, ChangeSetDocument } from './schemas/change-set.schema';
import { CellChange, StructuralOp } from './types/change-set.types';
import { CheckpointRecord, RestoreResult } from './types/checkpoint.types';

export interface MaybeAutoCheckpointInput {
  workbookId?: string;
  conversationId: string;
  actions: Action[];
}

export interface CreateManualCheckpointInput {
  workbookId: string;
  conversationId: string;
  label?: string;
}

type ChangeSetLean = {
  changeSetId: string;
  timestamp: Date;
  beforeState: unknown;
  changes: unknown;
  actions: unknown;
  structuralOps?: unknown;
  irreversibleActionTypes?: string[];
};

@Injectable()
export class CheckpointService {
  private readonly logger = new Logger(CheckpointService.name);

  constructor(
    @InjectModel(Checkpoint.name)
    private readonly checkpointModel: Model<CheckpointDocument>,
    @InjectModel(ChangeSet.name)
    private readonly changeSetModel: Model<ChangeSetDocument>,
    private readonly workflowTrace: WorkflowTraceService,
  ) {}

  /**
   * TASKS.md #27 — before persisting a change set that contains a destructive
   * action, snapshot a restore point anchored at the last already-applied
   * change set for this workbook, so the never-destroy-user-work principle
   * doesn't depend on the user having manually checkpointed first. Fails
   * open: a checkpointing failure must never block the user's real request.
   */
  async maybeAutoCheckpoint(input: MaybeAutoCheckpointInput): Promise<void> {
    if (!input.workbookId) return;
    const destructiveTypes = [
      ...new Set(input.actions.filter((a) => isDestructiveActionType(a.type)).map((a) => a.type)),
    ];
    if (destructiveTypes.length === 0) return;

    try {
      const anchor = await this.findLatestApplied(input.workbookId);
      await this.checkpointModel.create({
        checkpointId: randomUUID(),
        workbookId: input.workbookId,
        conversationId: input.conversationId,
        label: `Auto-checkpoint before ${destructiveTypes.join(', ')}`,
        trigger: 'auto',
        anchorChangeSetId: anchor?.changeSetId ?? '',
        createdAt: new Date(),
        status: 'active',
      });
    } catch (err: unknown) {
      this.warn('maybeAutoCheckpoint', err);
    }
  }

  /** TASKS.md #28 — explicit user-requested checkpoint. */
  async createManual(input: CreateManualCheckpointInput): Promise<CheckpointRecord> {
    const anchor = await this.findLatestApplied(input.workbookId);
    const doc = await this.checkpointModel.create({
      checkpointId: randomUUID(),
      workbookId: input.workbookId,
      conversationId: input.conversationId,
      label: input.label?.trim() || `Manual checkpoint ${new Date().toLocaleString()}`,
      trigger: 'manual',
      anchorChangeSetId: anchor?.changeSetId ?? '',
      createdAt: new Date(),
      status: 'active',
    });
    this.logger.log(`Checkpoint ${doc.checkpointId} created (manual) for workbook ${input.workbookId}`);
    return this.toRecord(doc);
  }

  async listByWorkbook(workbookId: string): Promise<CheckpointRecord[]> {
    const docs = await this.checkpointModel
      .find({ workbookId })
      .sort({ createdAt: -1 })
      .limit(50)
      .exec();
    return docs.map((doc) => this.toRecord(doc));
  }

  /**
   * TASKS.md #29 — restore algorithm, per DATABASE_SCHEMA.md §6.2's four steps:
   * find every applied change set for this workbook after the anchor, newest
   * first; build each one's full inverse; fail closed (no writes at all) if
   * any step in the chain can't be safely inverted; on success, mark every
   * reverted change set + the checkpoint itself, and write an audit trail.
   */
  async restore(checkpointId: string): Promise<RestoreResult> {
    const checkpoint = await this.checkpointModel.findOne({ checkpointId }).exec();
    if (!checkpoint) {
      throw new NotFoundException(`Checkpoint ${checkpointId} not found`);
    }

    // Ordered by _id, not timestamp: two change sets created in rapid succession
    // (e.g. a fast agentic loop) can share the same millisecond-resolution Date,
    // which would make a strict `timestamp: {$gt}` filter silently skip one of
    // them. MongoDB ObjectIds are strictly monotonic per collection regardless.
    const anchorObjectId = await this.resolveAnchorObjectId(
      checkpointId,
      checkpoint.anchorChangeSetId,
    );

    const chain = await this.changeSetModel
      .find({
        workbookId: checkpoint.workbookId,
        status: 'applied',
        ...(anchorObjectId ? ({ _id: { $gt: anchorObjectId } } as Record<string, unknown>) : {}),
      })
      .sort({ _id: -1 })
      .exec();

    if (chain.length === 0) {
      checkpoint.status = 'restored';
      checkpoint.restoredAt = new Date();
      await checkpoint.save();
      this.writeRestoreTrace(checkpoint, []);
      return { checkpoint: this.toRecord(checkpoint), revertedChangeSetIds: [], inverseActions: [] };
    }

    // Phase 1 — fail closed up front on any change set with a known-irreversible
    // action type (TASKS.md #18's catalog), before attempting anything.
    for (const doc of chain) {
      const irreversible = doc.irreversibleActionTypes ?? [];
      if (irreversible.length > 0) {
        throw new RestoreVerificationError(
          checkpointId,
          doc.changeSetId,
          `contains irreversible action type(s): ${irreversible.join(', ')}`,
        );
      }
    }

    // Phase 2 — dry-run verify every step's inverse (same self-verification
    // ChangeSetService.revert uses per change set) BEFORE writing anything to
    // the database, so a mid-chain failure never leaves a partial restore.
    const planned: { doc: ChangeSetDocument; inverseActions: Action[] }[] = [];
    for (const doc of chain) {
      const beforeState = doc.beforeState as Record<
        string,
        { value: unknown; formula: string; format: string }
      >;
      const changes = doc.changes as unknown as CellChange[];
      const cellInverseActions = beforeStateToInverseActions(beforeState, changes);
      const structuralOps = (doc.structuralOps ?? []) as unknown as StructuralOp[];
      const { pre, post } = structuralOpsToInverseActions(structuralOps);
      const inverseActions = [...pre, ...cellInverseActions, ...post];

      const expectedShadow = shadowFromBeforeState(beforeState);
      const originalActions = doc.actions as unknown as Action[];
      const currentShadow = virtualApply(expectedShadow, originalActions);
      const revertedShadow = virtualApply(currentShadow, inverseActions);
      const blockingChanges = diffShadowsFully(expectedShadow, revertedShadow);
      if (blockingChanges.length > 0) {
        throw new RestoreVerificationError(
          checkpointId,
          doc.changeSetId,
          `revert would not converge (${blockingChanges.length} cell(s))`,
        );
      }
      planned.push({ doc, inverseActions });
    }

    // Phase 3 — every step verified; commit newest-first, then the checkpoint.
    const revertedChangeSetIds: string[] = [];
    const allInverseActions: Action[] = [];
    for (const { doc, inverseActions } of planned) {
      doc.status = 'reverted';
      doc.revertedAt = new Date();
      await doc.save();
      revertedChangeSetIds.push(doc.changeSetId);
      allInverseActions.push(...inverseActions);
    }

    checkpoint.status = 'restored';
    checkpoint.restoredAt = new Date();
    await checkpoint.save();

    this.logger.log(
      `Checkpoint ${checkpointId} restored — reverted ${revertedChangeSetIds.length} change set(s)`,
    );
    this.writeRestoreTrace(checkpoint, revertedChangeSetIds);

    return {
      checkpoint: this.toRecord(checkpoint),
      revertedChangeSetIds,
      inverseActions: allInverseActions,
    };
  }

  private async findLatestApplied(workbookId: string): Promise<ChangeSetLean | null> {
    return this.changeSetModel
      .findOne({ workbookId, status: 'applied' }, { changeSetId: 1, timestamp: 1 })
      .sort({ timestamp: -1 })
      .lean()
      .exec() as unknown as Promise<ChangeSetLean | null>;
  }

  private async resolveAnchorObjectId(
    checkpointId: string,
    anchorChangeSetId: string,
  ): Promise<unknown | null> {
    if (!anchorChangeSetId) {
      return null;
    }
    const anchorDoc = await this.changeSetModel
      .findOne({ changeSetId: anchorChangeSetId }, { _id: 1 })
      .lean()
      .exec();
    if (!anchorDoc) {
      throw new RestoreVerificationError(
        checkpointId,
        anchorChangeSetId,
        'anchor change set no longer exists',
      );
    }
    return (anchorDoc as unknown as { _id: unknown })._id;
  }

  /** Fire-and-forget, mirroring WorkflowTraceService's own convention elsewhere. */
  private writeRestoreTrace(checkpoint: CheckpointDocument, revertedChangeSetIds: string[]): void {
    const traceId = `restore:${checkpoint.checkpointId}:${Date.now()}`;
    this.workflowTrace.startTrace({
      traceId,
      conversationId: checkpoint.conversationId,
      workbookId: checkpoint.workbookId,
      message: `Restore checkpoint ${checkpoint.checkpointId} (${checkpoint.label})`,
      mode: 'action',
    });
    this.workflowTrace.appendNode(traceId, {
      id: 'restore',
      type: 'restore',
      label: 'Checkpoint Restore',
      status: 'success',
      input: { checkpointId: checkpoint.checkpointId, anchorChangeSetId: checkpoint.anchorChangeSetId },
      output: { revertedChangeSetIds },
    });
    this.workflowTrace.finalize(traceId, { status: 'completed', changeSetId: checkpoint.anchorChangeSetId || undefined });
  }

  private warn(op: string, err: unknown): void {
    const msg = err instanceof Error ? err.message : String(err);
    this.logger.warn(`CheckpointService.${op} failed: ${msg}`);
  }

  private toRecord(doc: CheckpointDocument): CheckpointRecord {
    return {
      checkpointId: doc.checkpointId,
      workbookId: doc.workbookId,
      conversationId: doc.conversationId,
      label: doc.label,
      trigger: doc.trigger,
      anchorChangeSetId: doc.anchorChangeSetId,
      createdAt: doc.createdAt,
      status: doc.status,
      restoredAt: doc.restoredAt,
    };
  }
}
