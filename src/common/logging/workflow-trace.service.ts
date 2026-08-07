import { Injectable, Logger } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { sanitizeLogBody } from './log-body.util';
import {
  WorkflowEdge,
  WorkflowNode,
  WorkflowNodeStatus,
  WorkflowNodeType,
  WorkflowTrace,
  WorkflowTraceStatus,
} from './schemas/workflow-trace.schema';

const MAX_PAYLOAD_CHARS = 4096;

export interface StartWorkflowTraceInput {
  traceId: string;
  conversationId?: string;
  message: string;
  mode?: string;
  /** Raw request body / sheet summary — sanitized before store. */
  requestInput?: unknown;
}

export interface AppendWorkflowNodeInput {
  id: string;
  type: WorkflowNodeType;
  label: string;
  status: WorkflowNodeStatus;
  startedAt?: Date;
  endedAt?: Date;
  durationMs?: number;
  input?: unknown;
  output?: unknown;
  meta?: Record<string, unknown>;
  /** Explicit edge source; defaults to previous lastNodeId. */
  linkFrom?: string | null;
  /** Skip creating an edge (e.g. first node after start already linked). */
  skipEdge?: boolean;
}

export interface FinalizeWorkflowTraceInput {
  status: WorkflowTraceStatus;
  durationMs?: number;
  changeSetId?: string;
  route?: string;
  tier?: number;
  sseOutput?: unknown;
}

@Injectable()
export class WorkflowTraceService {
  private readonly logger = new Logger(WorkflowTraceService.name);

  constructor(
    @InjectModel(WorkflowTrace.name)
    private readonly workflowTraceModel: Model<WorkflowTrace>,
  ) {}

  /** Fire-and-forget create of a running trace with frontend_in node. */
  startTrace(input: StartWorkflowTraceInput): void {
    void this.startTraceAsync(input).catch((err: unknown) => {
      this.warn('startTrace', err);
    });
  }

  /** Fire-and-forget append of a graph node (+ optional edge). */
  appendNode(traceId: string, node: AppendWorkflowNodeInput): void {
    if (!traceId || traceId === '-') return;
    void this.appendNodeAsync(traceId, node).catch((err: unknown) => {
      this.warn('appendNode', err);
    });
  }

  /** Fire-and-forget finalize (status + optional sse_out node). */
  finalize(traceId: string, input: FinalizeWorkflowTraceInput): void {
    if (!traceId || traceId === '-') return;
    void this.finalizeAsync(traceId, input).catch((err: unknown) => {
      this.warn('finalize', err);
    });
  }

  setMeta(
    traceId: string,
    meta: { route?: string; tier?: number; changeSetId?: string },
  ): void {
    if (!traceId || traceId === '-') return;
    void this.workflowTraceModel
      .updateOne(
        { traceId },
        {
          $set: {
            ...(meta.route !== undefined ? { route: meta.route } : {}),
            ...(meta.tier !== undefined ? { tier: meta.tier } : {}),
            ...(meta.changeSetId !== undefined ? { changeSetId: meta.changeSetId } : {}),
          },
        },
      )
      .exec()
      .catch((err: unknown) => this.warn('setMeta', err));
  }

  /** Append terminal preview/accept/reject node keyed by changeSetId. */
  appendTerminalByChangeSet(
    changeSetId: string,
    node: AppendWorkflowNodeInput,
    status?: WorkflowTraceStatus,
  ): void {
    if (!changeSetId) return;
    void this.appendTerminalAsync({ changeSetId }, node, status).catch((err: unknown) => {
      this.warn('appendTerminalByChangeSet', err);
    });
  }

  /** Append terminal node keyed by conversationId (most recent running/awaiting trace). */
  appendTerminalByConversationId(
    conversationId: string,
    node: AppendWorkflowNodeInput,
    status?: WorkflowTraceStatus,
  ): void {
    if (!conversationId) return;
    void this.appendTerminalAsync({ conversationId }, node, status).catch((err: unknown) => {
      this.warn('appendTerminalByConversationId', err);
    });
  }

  sanitizePayload(value: unknown): unknown {
    return truncatePayload(sanitizeLogBody(value) ?? value);
  }

  private async startTraceAsync(input: StartWorkflowTraceInput): Promise<void> {
    const now = new Date();
    const frontendNode: WorkflowNode = {
      id: 'frontend_in',
      type: 'frontend_in',
      label: 'Frontend Request',
      status: 'success',
      startedAt: now,
      endedAt: now,
      durationMs: 0,
      input: this.sanitizePayload(input.requestInput ?? {
        message: input.message,
        mode: input.mode,
        conversationId: input.conversationId,
      }),
      output: {
        conversationId: input.conversationId,
        traceId: input.traceId,
      },
    };

    await this.workflowTraceModel.findOneAndUpdate(
      { traceId: input.traceId },
      {
        $setOnInsert: {
          ts: now,
          traceId: input.traceId,
          conversationId: input.conversationId,
          message: input.message.slice(0, 2000),
          mode: input.mode,
          status: 'running' as WorkflowTraceStatus,
          nodes: [frontendNode],
          edges: [],
          lastNodeId: frontendNode.id,
        },
      },
      { upsert: true },
    );
  }

  private async appendNodeAsync(
    traceId: string,
    nodeInput: AppendWorkflowNodeInput,
  ): Promise<void> {
    const now = new Date();
    const node: WorkflowNode = {
      id: nodeInput.id,
      type: nodeInput.type,
      label: nodeInput.label,
      status: nodeInput.status,
      startedAt: nodeInput.startedAt,
      endedAt: nodeInput.endedAt ?? now,
      durationMs: nodeInput.durationMs,
      input: nodeInput.input !== undefined ? this.sanitizePayload(nodeInput.input) : undefined,
      output: nodeInput.output !== undefined ? this.sanitizePayload(nodeInput.output) : undefined,
      meta: nodeInput.meta,
    };

    const existing = await this.workflowTraceModel.findOne({ traceId }).lean().exec();
    if (!existing) {
      this.logger.debug(`appendNode: no trace for ${traceId}, skipping ${node.id}`);
      return;
    }

    // Idempotent: update existing node in place if same id already present.
    const already = (existing.nodes ?? []).some((n) => n.id === node.id);
    if (already) {
      await this.workflowTraceModel.updateOne(
        { traceId, 'nodes.id': node.id },
        {
          $set: {
            'nodes.$': node,
            lastNodeId: node.id,
          },
        },
      );
      return;
    }

    const linkFrom =
      nodeInput.skipEdge || nodeInput.linkFrom === null
        ? null
        : (nodeInput.linkFrom ?? existing.lastNodeId ?? null);

    const edge: WorkflowEdge | null =
      linkFrom && linkFrom !== node.id
        ? { id: `e_${linkFrom}_${node.id}`, source: linkFrom, target: node.id }
        : null;

    await this.workflowTraceModel.updateOne(
      { traceId },
      {
        $push: {
          nodes: node,
          ...(edge ? { edges: edge } : {}),
        },
        $set: { lastNodeId: node.id },
      },
    );
  }

  private async finalizeAsync(
    traceId: string,
    input: FinalizeWorkflowTraceInput,
  ): Promise<void> {
    const existing = await this.workflowTraceModel.findOne({ traceId }).lean().exec();
    if (!existing) return;

    if (input.sseOutput !== undefined) {
      await this.appendNodeAsync(traceId, {
        id: 'sse_out',
        type: 'sse_out',
        label: 'SSE → Frontend',
        status: input.status === 'failed' ? 'failed' : 'success',
        input: {
          route: input.route ?? existing.route,
          tier: input.tier ?? existing.tier,
          changeSetId: input.changeSetId ?? existing.changeSetId,
        },
        output: input.sseOutput,
        linkFrom: existing.lastNodeId,
      });
    }

    await this.workflowTraceModel.updateOne(
      { traceId },
      {
        $set: {
          status: input.status,
          ...(input.durationMs !== undefined ? { durationMs: input.durationMs } : {}),
          ...(input.changeSetId !== undefined ? { changeSetId: input.changeSetId } : {}),
          ...(input.route !== undefined ? { route: input.route } : {}),
          ...(input.tier !== undefined ? { tier: input.tier } : {}),
        },
      },
    );
  }

  private async appendTerminalAsync(
    filter: { changeSetId?: string; conversationId?: string },
    nodeInput: AppendWorkflowNodeInput,
    status?: WorkflowTraceStatus,
  ): Promise<void> {
    const existing = filter.changeSetId
      ? await this.workflowTraceModel
          .findOne({ changeSetId: filter.changeSetId })
          .sort({ ts: -1 })
          .lean()
          .exec()
      : await this.workflowTraceModel
          .findOne({
            conversationId: filter.conversationId,
            status: {
              $in: [
                'running',
                'awaiting_accept',
                'completed',
                'clarifying',
              ] as WorkflowTraceStatus[],
            },
          })
          .sort({ ts: -1 })
          .lean()
          .exec();
    if (!existing) return;

    await this.appendNodeAsync(existing.traceId, {
      ...nodeInput,
      linkFrom: nodeInput.linkFrom ?? existing.lastNodeId,
    });

    if (status) {
      await this.workflowTraceModel.updateOne(
        { _id: existing._id },
        { $set: { status } },
      );
    }
  }

  private warn(op: string, err: unknown): void {
    const msg = err instanceof Error ? err.message : String(err);
    this.logger.warn(`WorkflowTraceService.${op} failed: ${msg}`);
  }
}

function truncatePayload(value: unknown): unknown {
  if (value === undefined || value === null) return value;
  if (typeof value === 'string') {
    return value.length <= MAX_PAYLOAD_CHARS
      ? value
      : `${value.slice(0, MAX_PAYLOAD_CHARS)}… [truncated ${value.length - MAX_PAYLOAD_CHARS} chars]`;
  }
  try {
    const json = JSON.stringify(value);
    if (json.length <= MAX_PAYLOAD_CHARS) return value;
    return {
      _truncated: true,
      preview: `${json.slice(0, MAX_PAYLOAD_CHARS)}…`,
      originalChars: json.length,
    };
  } catch {
    return '[unserializable]';
  }
}
