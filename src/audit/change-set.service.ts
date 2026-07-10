import { Injectable, Logger, NotFoundException } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { randomUUID } from 'crypto';
import { Model } from 'mongoose';
import { Action } from '../agents/types/agent.types';
import { WorkbookContext } from '../agents/types/agent.types';
import { buildShadowWorkbook } from '../virtual/shadowWorkbook';
import { virtualApply } from '../virtual/virtualApply';
import {
  beforeStateToInverseActions,
  generateDiff,
  snapshotBeforeState,
} from './diff.engine';
import { ChangeSet, ChangeSetDocument } from './schemas/change-set.schema';
import { CellChange, ChangeSetRecord } from './types/change-set.types';

export interface CreatePreviewInput {
  conversationId: string;
  traceId: string;
  prompt: string;
  context: WorkbookContext;
  actions: Action[];
}

@Injectable()
export class ChangeSetService {
  private readonly logger = new Logger(ChangeSetService.name);

  constructor(
    @InjectModel(ChangeSet.name)
    private readonly changeSetModel: Model<ChangeSetDocument>,
  ) {}

  async createPreview(input: CreatePreviewInput): Promise<ChangeSetRecord> {
    const beforeShadow = buildShadowWorkbook(input.context);
    const beforeState = snapshotBeforeState(beforeShadow);
    const afterShadow = virtualApply(beforeShadow, input.actions);
    const changes = generateDiff(beforeShadow, afterShadow);
    const changeSetId = randomUUID();

    const doc = await this.changeSetModel.create({
      changeSetId,
      conversationId: input.conversationId,
      traceId: input.traceId,
      timestamp: new Date(),
      prompt: input.prompt,
      beforeState,
      changes,
      actions: input.actions as unknown as Record<string, unknown>[],
      status: 'previewed',
    });

    this.logger.log(
      `Change set ${changeSetId} previewed: ${changes.length} cell(s) for conversation ${input.conversationId}`,
    );

    return this.toRecord(doc);
  }

  async markApplied(changeSetId: string): Promise<ChangeSetRecord> {
    const existing = await this.changeSetModel.findOne({ changeSetId }).exec();
    if (existing?.status === 'applied') {
      return this.toRecord(existing);
    }

    const doc = await this.changeSetModel.findOneAndUpdate(
      { changeSetId, status: 'previewed' },
      { status: 'applied', appliedAt: new Date() },
      { new: true },
    );
    if (!doc) {
      throw new NotFoundException(`Change set ${changeSetId} not found or not previewed`);
    }
    this.logger.log(`Change set ${changeSetId} marked applied`);
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
    const inverseActions = beforeStateToInverseActions(beforeState, changes);

    doc.status = 'reverted';
    doc.revertedAt = new Date();
    await doc.save();

    this.logger.log(`Change set ${changeSetId} reverted with ${inverseActions.length} inverse action(s)`);
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

  private toRecord(doc: ChangeSetDocument): ChangeSetRecord {
    return {
      changeSetId: doc.changeSetId,
      conversationId: doc.conversationId,
      traceId: doc.traceId,
      timestamp: doc.timestamp,
      prompt: doc.prompt,
      beforeState: doc.beforeState as ChangeSetRecord['beforeState'],
      changes: doc.changes as CellChange[],
      actions: doc.actions as unknown as Action[],
      status: doc.status as ChangeSetRecord['status'],
      appliedAt: doc.appliedAt,
      revertedAt: doc.revertedAt,
    };
  }
}
