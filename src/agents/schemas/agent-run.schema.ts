import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes } from 'mongoose';
import type { Action, SubTask, WorkbookContext } from '../types/agent.types';

export type AgentRunDocument = HydratedDocument<AgentRun>;

/**
 * A stepwise run is working state, not an audit record — the `change_sets` each
 * wave produces are the durable artifact and carry their own retention. 24h
 * matches the "scratch buffer for resuming today's thread" role, and means a
 * run nobody continues cannot live forever. STEPWISE_EXECUTION.md §4.
 */
export const AGENT_RUN_TTL_MS =
  Number(process.env.AGENT_RUN_TTL_HOURS ?? 24) * 60 * 60 * 1000;

/** Per-subtask outcome, carried across requests. STEPWISE_EXECUTION.md §4. */
@Schema({ _id: false })
export class AgentRunSubtaskState {
  @Prop({ type: String, required: true })
  subtaskId!: string;

  @Prop({ type: [SchemaTypes.Mixed], default: [] })
  actions!: Action[];

  @Prop({ type: Boolean, default: false })
  completed!: boolean;

  @Prop({ type: Boolean, required: false })
  verified?: boolean;

  @Prop({ type: String, required: false })
  failedReason?: string;

  /**
   * The user's call on the wave this subtask belonged to. `skipped` covers both
   * a cascade-skip (a dependency was rejected) and a deterministic-check failure
   * the run chose to continue past — STEPWISE_EXECUTION.md SD-4.
   */
  @Prop({ type: String, enum: ['accepted', 'rejected', 'skipped'], required: false })
  decision?: 'accepted' | 'rejected' | 'skipped';
}

export const AgentRunSubtaskStateSchema =
  SchemaFactory.createForClass(AgentRunSubtaskState);

@Schema({ timestamps: true, collection: 'agent_runs' })
export class AgentRun {
  @Prop({ type: String, required: true, unique: true, index: true })
  runId!: string;

  @Prop({ type: String, required: true, index: true })
  conversationId!: string;

  /**
   * Ownership check for `/continue` — resolved server-side from the session,
   * never read from the request body, same discipline as
   * `Conversation.userId` (TASKS.md #170). Without this any caller holding a
   * runId could drive someone else's build.
   */
  @Prop({ type: String, required: false, index: true })
  userId?: string;

  @Prop({ type: String, required: true })
  traceId!: string;

  @Prop({ type: String, required: true })
  prompt!: string;

  @Prop({
    type: String,
    enum: ['awaiting_decision', 'running', 'completed', 'failed', 'abandoned'],
    default: 'awaiting_decision',
    index: true,
  })
  status!: 'awaiting_decision' | 'running' | 'completed' | 'failed' | 'abandoned';

  /** Index of the wave most recently emitted; -1 before the first one. */
  @Prop({ type: Number, default: -1 })
  waveIndex!: number;

  @Prop({ type: Number, required: true })
  waveTotal!: number;

  /** The whole plan, frozen at plan time — SD-2 never re-plans mid-run. */
  @Prop({ type: [SchemaTypes.Mixed], default: [] })
  subtasks!: SubTask[];

  /** Subtask ids per wave, frozen alongside `subtasks`. */
  @Prop({ type: [[String]], default: [] })
  waves!: string[][];

  @Prop({ type: [AgentRunSubtaskStateSchema], default: [] })
  subtaskStates!: AgentRunSubtaskState[];

  /** Base workbook context, refreshed from each `/continue` readback. */
  @Prop({ type: SchemaTypes.Mixed, required: true })
  context!: WorkbookContext;

  /** One per emitted wave — seeds the next wave's `dependsOnChangeSetId`. */
  @Prop({ type: [String], default: [] })
  changeSetIds!: string[];

  @Prop({ type: String, required: false })
  promptContext?: string;

  @Prop({ type: [SchemaTypes.Mixed], default: [] })
  conversationHistory!: { role: 'user' | 'assistant'; content: string }[];

  @Prop({ type: String, required: false })
  routerAssumption?: string;

  @Prop({ type: Date, default: () => new Date(Date.now() + AGENT_RUN_TTL_MS) })
  expiresAt!: Date;
}

export const AgentRunSchema = SchemaFactory.createForClass(AgentRun);
AgentRunSchema.index({ expiresAt: 1 }, { expireAfterSeconds: 0 });
