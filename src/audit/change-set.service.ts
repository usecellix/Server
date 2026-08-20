import { Injectable, Logger, NotFoundException, Optional } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { randomUUID } from 'crypto';
import { Model } from 'mongoose';
import { Action } from '../agents/types/agent.types';
import { WorkbookContext } from '../agents/types/agent.types';
import { WorkflowTraceService } from '../common/logging/workflow-trace.service';
import { Conversation, ConversationDocument } from '../excel-ai/schemas/conversation.schema';
import { buildShadowWorkbook } from '../virtual/shadowWorkbook';
import { virtualApply } from '../virtual/virtualApply';
import { CheckpointService } from './checkpoint.service';
import {
  beforeStateToInverseActions,
  captureStructuralOps,
  computeUnintendedChanges,
  detectIntroducedFormulaErrors,
  diffShadowsFully,
  excludeStructurallyOwnedChanges,
  generateDiff,
  shadowFromBeforeState,
  snapshotBeforeState,
  structuralOpsToInverseActions,
} from './diff.engine';
import { RevertVerificationError } from './errors/revert-verification.error';
import { computeIrreversibleActionTypes } from './reversibility-catalog';
import { ChangeSet, ChangeSetDocument } from './schemas/change-set.schema';
import {
  CellChange,
  ChangeSetRecord,
  FormulaErrorChange,
  StructuralOp,
} from './types/change-set.types';
import {
  assertDomainToolProvenance,
  ProvenanceContext,
} from './utils/provenance.util';

export interface CreatePreviewInput {
  conversationId: string;
  traceId: string;
  prompt: string;
  context: WorkbookContext;
  actions: Action[];
  /** Optional citations / exception flags threaded from Tier 2/3 or domain tools */
  provenance?: ProvenanceContext;
  /**
   * Durable per-workbook identity (TASKS.md #24). Optional explicit override —
   * mainly for tests. Production callers don't need to pass this: createPreview
   * resolves it from the originating conversation by conversationId, so every
   * call site automatically gets it right rather than needing to remember to
   * thread it through the (many, deep) call chains that lead here — the same
   * "declared but not wired up" failure class action-catalog.ts guards against
   * for action types.
   */
  workbookId?: string;
}

@Injectable()
export class ChangeSetService {
  private readonly logger = new Logger(ChangeSetService.name);

  constructor(
    @InjectModel(ChangeSet.name)
    private readonly changeSetModel: Model<ChangeSetDocument>,
    @InjectModel(Conversation.name)
    private readonly conversationModel: Model<ConversationDocument>,
    private readonly workflowTrace: WorkflowTraceService,
    // Optional so existing hand-rolled test constructions (new ChangeSetService(a,b,c))
    // keep working unmodified — a missing CheckpointService just means auto-checkpointing
    // is skipped, never a hard failure of the actual write path (TASKS.md #27).
    @Optional()
    private readonly checkpointService?: CheckpointService,
  ) {}

  async createPreview(input: CreatePreviewInput): Promise<ChangeSetRecord> {
    assertDomainToolProvenance(input.provenance);

    const beforeShadow = buildShadowWorkbook(input.context);
    const beforeState = snapshotBeforeState(beforeShadow);
    const afterShadow = virtualApply(beforeShadow, input.actions);
    const structuralOps = captureStructuralOps(beforeShadow, afterShadow, input.actions);
    const rawChanges = generateDiff(beforeShadow, afterShadow);
    // Cells a structural column op's own inverse already fully restores via its shift —
    // the generic per-cell inverse must not also touch them (see TASKS.md #13).
    const baseChanges = excludeStructurallyOwnedChanges(rawChanges, structuralOps);
    const changes = this.attachProvenance(baseChanges, input.provenance);
    // Full action objects, not just .type — TASKS.md #40's CONDITIONAL_FORMAT special case
    // needs to see `existingRuleId` to distinguish a revertible create from an unrevertible modify.
    const irreversibleActionTypes = computeIrreversibleActionTypes(input.actions);
    // PRD Tier A metrics (TASKS.md #48/#49) — computed from the same before/after shadow
    // pair and diff this method already builds, so it's tier-agnostic (Tier 0-3 alike),
    // not dependent on Tier 3's planner/agenticLoop internals.
    const unintendedChanges = computeUnintendedChanges(changes, input.actions);
    const formulaErrorsIntroduced = detectIntroducedFormulaErrors(changes);
    const changeSetId = randomUUID();
    const workbookId = input.workbookId ?? (await this.resolveWorkbookId(input.conversationId));

    // TASKS.md #27 — before persisting a change set containing a destructive
    // action, snapshot a restore point anchored at the last applied change set,
    // so never-destroy-user-work doesn't depend on the user checkpointing first.
    await this.checkpointService?.maybeAutoCheckpoint({
      workbookId,
      conversationId: input.conversationId,
      actions: input.actions,
    });

    const doc = await this.changeSetModel.create({
      changeSetId,
      conversationId: input.conversationId,
      ...(workbookId ? { workbookId } : {}),
      traceId: input.traceId,
      timestamp: new Date(),
      prompt: input.prompt,
      beforeState,
      changes: changes as unknown as ChangeSetDocument['changes'],
      actions: input.actions as unknown as Record<string, unknown>[],
      structuralOps: structuralOps as unknown as ChangeSetDocument['structuralOps'],
      irreversibleActionTypes,
      unintendedChanges: unintendedChanges as unknown as ChangeSetDocument['unintendedChanges'],
      formulaErrorsIntroduced:
        formulaErrorsIntroduced as unknown as ChangeSetDocument['formulaErrorsIntroduced'],
      status: 'previewed',
      provenanceConfidence: input.provenance?.confidence,
    });

    this.logger.log(
      `Change set ${changeSetId} previewed: ${changes.length} cell(s) for conversation ${input.conversationId}` +
        (input.provenance?.sourceRefs?.length
          ? ` sourceRefs=${input.provenance.sourceRefs.length}`
          : '') +
        (unintendedChanges.length ? ` unintendedChanges=${unintendedChanges.length}` : '') +
        (formulaErrorsIntroduced.length
          ? ` formulaErrorsIntroduced=${formulaErrorsIntroduced.length}`
          : ''),
    );

    return this.toRecord(doc);
  }

  /**
   * TASKS.md #40's apply-endpoint contract change, extended at #15 to CREATE_CHART. A
   * CONDITIONAL_FORMAT/CREATE_CHART structuralOp is captured at preview time with no
   * runtime id (Office.js doesn't assign a rule id / chart name until the real apply
   * happens); the frontend reads the real id back right after creating each one for
   * real and reports it here, matched to its structuralOp entry by sheetName+range
   * (CONDITIONAL_FORMAT) or sheetName+sourceRange (CREATE_CHART), before the change set
   * is marked applied. Unmatched entries (e.g. one the frontend failed to apply, or a
   * batch with no such actions at all) are left untouched — structuralOpsToInverseActions
   * fails closed on a still-missing id at revert time rather than here.
   */
  async markApplied(
    changeSetId: string,
    createdConditionalFormatIds?: { sheetName: string; range: string; ruleId: string }[],
    createdChartIds?: { sheetName: string; sourceRange: string; chartId: string }[],
  ): Promise<ChangeSetRecord> {
    const existing = await this.changeSetModel.findOne({ changeSetId }).exec();
    if (existing?.status === 'applied') {
      return this.toRecord(existing);
    }

    const update: Record<string, unknown> = { status: 'applied', appliedAt: new Date() };

    if (
      existing &&
      ((createdConditionalFormatIds && createdConditionalFormatIds.length > 0) ||
        (createdChartIds && createdChartIds.length > 0))
    ) {
      const structuralOps = (existing.structuralOps ?? []) as unknown as StructuralOp[];
      update.structuralOps = structuralOps.map((op) => {
        if (op.opType === 'CONDITIONAL_FORMAT') {
          const match = createdConditionalFormatIds?.find(
            (created) => created.sheetName === op.sheetName && created.range === op.params.range,
          );
          return match ? { ...op, params: { ...op.params, ruleId: match.ruleId } } : op;
        }
        if (op.opType === 'CREATE_CHART') {
          const match = createdChartIds?.find(
            (created) =>
              created.sheetName === op.sheetName && created.sourceRange === op.params.sourceRange,
          );
          return match ? { ...op, params: { ...op.params, chartId: match.chartId } } : op;
        }
        return op;
      });
    }

    const doc = await this.changeSetModel.findOneAndUpdate(
      { changeSetId, status: 'previewed' },
      update,
      { new: true },
    );
    if (!doc) {
      throw new NotFoundException(`Change set ${changeSetId} not found or not previewed`);
    }
    this.logger.log(`Change set ${changeSetId} marked applied`);
    this.workflowTrace.appendTerminalByChangeSet(
      changeSetId,
      {
        id: `accept:${Date.now()}`,
        type: 'accept',
        label: 'Accept / Apply',
        status: 'success',
        input: { changeSetId },
        output: { status: 'applied', appliedAt: doc.appliedAt },
      },
      'accepted',
    );
    return this.toRecord(doc);
  }

  async revert(changeSetId: string): Promise<{ changeSet: ChangeSetRecord; inverseActions: Action[] }> {
    const doc = await this.changeSetModel.findOne({ changeSetId });
    if (!doc) {
      throw new NotFoundException(`Change set ${changeSetId} not found`);
    }
    if (doc.status !== 'applied') {
      throw new NotFoundException(
        `Change set ${changeSetId} cannot be reverted (status: ${doc.status})`,
      );
    }

    const beforeState = doc.beforeState as Record<string, { value: unknown; formula: string; format: string }>;
    const changes = doc.changes as CellChange[];
    const cellInverseActions = beforeStateToInverseActions(beforeState, changes);
    const structuralOps = (doc.structuralOps ?? []) as unknown as StructuralOp[];
    const { pre, post } = structuralOpsToInverseActions(structuralOps);
    const inverseActions = [...pre, ...cellInverseActions, ...post];

    // Fail-closed self-verification (TASKS.md #19): dry-run the inverse against a shadow
    // rebuilt from beforeState + the original forward actions, and refuse the revert
    // entirely — never partially — if the result doesn't converge back to beforeState.
    const expectedShadow = shadowFromBeforeState(beforeState);
    const originalActions = doc.actions as unknown as Action[];
    const currentShadow = virtualApply(expectedShadow, originalActions);
    const revertedShadow = virtualApply(currentShadow, inverseActions);
    const blockingChanges = diffShadowsFully(expectedShadow, revertedShadow);
    if (blockingChanges.length > 0) {
      this.logger.warn(
        `Change set ${changeSetId} revert refused — ${blockingChanges.length} cell(s) would not converge`,
      );
      throw new RevertVerificationError(changeSetId, blockingChanges);
    }

    doc.status = 'reverted';
    doc.revertedAt = new Date();
    await doc.save();

    this.logger.log(`Change set ${changeSetId} reverted with ${inverseActions.length} inverse action(s)`);
    this.workflowTrace.appendTerminalByChangeSet(
      changeSetId,
      {
        id: `reject:${Date.now()}`,
        type: 'reject',
        label: 'Revert',
        status: 'success',
        input: { changeSetId },
        output: { status: 'reverted', inverseActionCount: inverseActions.length },
      },
      'rejected',
    );
    return { changeSet: this.toRecord(doc), inverseActions };
  }

  async getHistory(conversationId: string): Promise<ChangeSetRecord[]> {
    const docs = await this.changeSetModel
      .find({ conversationId })
      .sort({ timestamp: -1 })
      .limit(50)
      .exec();
    return docs.map((doc) => this.toRecord(doc));
  }

  async getById(changeSetId: string): Promise<ChangeSetRecord | null> {
    const doc = await this.changeSetModel.findOne({ changeSetId }).exec();
    return doc ? this.toRecord(doc) : null;
  }

  async getByDateRange(fromDate: Date, toDate: Date): Promise<ChangeSetRecord[]> {
    const docs = await this.changeSetModel
      .find({ timestamp: { $gte: fromDate, $lte: toDate } })
      .sort({ timestamp: -1 })
      .limit(5000)
      .exec();
    return docs.map((doc) => this.toRecord(doc));
  }

  private attachProvenance(
    changes: CellChange[],
    provenance?: ProvenanceContext,
  ): CellChange[] {
    if (!provenance) return changes;
    const sourceRefs = provenance.sourceRefs;
    const exceptionFlags = provenance.exceptionFlags;
    if (!sourceRefs?.length && !exceptionFlags?.length) {
      return changes;
    }

    return changes.map((change) => ({
      ...change,
      ...(sourceRefs?.length ? { sourceRefs } : {}),
      ...(exceptionFlags?.length ? { exceptionFlags } : {}),
    }));
  }

  /** Best-effort lookup — a missing/expired conversation just means no workbookId, not an error. */
  private async resolveWorkbookId(conversationId: string): Promise<string | undefined> {
    try {
      const doc = await this.conversationModel
        .findOne({ conversationId }, { workbookId: 1 })
        .lean()
        .exec();
      return doc?.workbookId;
    } catch (err: unknown) {
      this.warn('resolveWorkbookId', err);
      return undefined;
    }
  }

  private warn(op: string, err: unknown): void {
    const msg = err instanceof Error ? err.message : String(err);
    this.logger.warn(`ChangeSetService.${op} failed: ${msg}`);
  }

  private toRecord(doc: ChangeSetDocument): ChangeSetRecord {
    return {
      changeSetId: doc.changeSetId,
      conversationId: doc.conversationId,
      workbookId: doc.workbookId,
      traceId: doc.traceId,
      timestamp: doc.timestamp,
      prompt: doc.prompt,
      beforeState: doc.beforeState as ChangeSetRecord['beforeState'],
      changes: doc.changes as CellChange[],
      actions: doc.actions as unknown as Action[],
      structuralOps: (doc.structuralOps ?? []) as ChangeSetRecord['structuralOps'],
      irreversibleActionTypes: doc.irreversibleActionTypes ?? [],
      unintendedChanges: (doc.unintendedChanges ?? []) as CellChange[],
      formulaErrorsIntroduced: (doc.formulaErrorsIntroduced ?? []) as FormulaErrorChange[],
      status: doc.status as ChangeSetRecord['status'],
      appliedAt: doc.appliedAt,
      revertedAt: doc.revertedAt,
      provenanceConfidence: (doc as ChangeSetDocument & { provenanceConfidence?: number })
        .provenanceConfidence,
    };
  }
}
