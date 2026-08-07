import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';
import { LOG_TTL_SECONDS } from './request-log.schema';

export type WorkflowTraceDocument = HydratedDocument<WorkflowTrace>;

export type WorkflowTraceStatus =
  | 'running'
  | 'completed'
  | 'failed'
  | 'clarifying'
  | 'awaiting_accept'
  | 'accepted'
  | 'rejected';

export type WorkflowNodeType =
  | 'frontend_in'
  | 'router'
  | 'tier'
  | 'planner'
  | 'executor'
  | 'verifier'
  | 'tool'
  | 'changeset'
  | 'sse_out'
  | 'preview'
  | 'accept'
  | 'reject'
  | 'error';

export type WorkflowNodeStatus = 'pending' | 'running' | 'success' | 'failed' | 'skipped';

export interface WorkflowNode {
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
}

export interface WorkflowEdge {
  id: string;
  source: string;
  target: string;
}

@Schema({
  collection: 'workflow_traces',
  versionKey: false,
})
export class WorkflowTrace {
  _id!: Types.ObjectId;

  /** TTL field — documents auto-delete 3 days after this timestamp. */
  @Prop({ type: Date, required: true })
  ts!: Date;

  @Prop({ type: String, required: true, index: true })
  traceId!: string;

  @Prop({ type: String, index: true })
  conversationId?: string;

  @Prop({ type: String, index: true })
  changeSetId?: string;

  @Prop({ type: String, required: true })
  message!: string;

  @Prop({ type: String })
  mode?: string;

  @Prop({ type: String })
  route?: string;

  @Prop({ type: Number })
  tier?: number;

  @Prop({ type: String, required: true, index: true })
  status!: WorkflowTraceStatus;

  @Prop({ type: Number })
  durationMs?: number;

  @Prop({ type: [SchemaTypes.Mixed], default: [] })
  nodes!: WorkflowNode[];

  @Prop({ type: [SchemaTypes.Mixed], default: [] })
  edges!: WorkflowEdge[];

  /** Last appended node id — used to auto-link edges. */
  @Prop({ type: String })
  lastNodeId?: string;
}

export const WorkflowTraceSchema = SchemaFactory.createForClass(WorkflowTrace);
WorkflowTraceSchema.index({ ts: 1 }, { expireAfterSeconds: LOG_TTL_SECONDS });
WorkflowTraceSchema.index({ conversationId: 1, ts: -1 });
WorkflowTraceSchema.index({ changeSetId: 1 });
