import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument, SchemaTypes, Types } from 'mongoose';

export type ChangeSetDocument = HydratedDocument<ChangeSet>;

@Schema({ _id: false })
export class CellSnapshotSchema {
  @Prop({ type: SchemaTypes.Mixed })
  value!: unknown;

  @Prop({ type: String, default: '' })
  formula!: string;

  @Prop({ type: String, default: 'General' })
  format!: string;
}

@Schema({ _id: false })
export class StructuralOpSchema {
  @Prop({ type: String, required: true })
  opType!: string;

  @Prop({ type: String, required: true })
  sheetName!: string;

  @Prop({ type: SchemaTypes.Mixed, default: {} })
  params!: Record<string, unknown>;

  @Prop({ type: Date, required: true, default: () => new Date() })
  appliedAt!: Date;
}

@Schema({ _id: false })
export class FormulaErrorChangeSchema {
  @Prop({ type: String, required: true })
  cell!: string;

  @Prop({ type: String, required: true })
  sheet!: string;

  @Prop({ type: String, required: true })
  error!: string;
}

@Schema({ _id: false })
export class CellChangeSchema {
  @Prop({ type: String, required: true })
  cell!: string;

  @Prop({ type: String, required: true })
  sheet!: string;

  @Prop({ type: SchemaTypes.Mixed })
  before!: unknown;

  @Prop({ type: SchemaTypes.Mixed })
  after!: unknown;

  @Prop({ type: String })
  formula?: string;

  @Prop({ type: Boolean, required: true })
  isHardcoded!: boolean;

  @Prop({ type: [SchemaTypes.Mixed], required: false })
  sourceRefs?: Record<string, unknown>[];

  @Prop({ type: [SchemaTypes.Mixed], required: false })
  exceptionFlags?: Record<string, unknown>[];
}

@Schema({
  collection: 'change_sets',
  versionKey: false,
})
export class ChangeSet {
  _id!: Types.ObjectId;

  @Prop({ type: String, required: true, unique: true, index: true })
  changeSetId!: string;

  @Prop({ type: String, required: true, index: true })
  conversationId!: string;

  /**
   * Durable per-workbook identity (TASKS.md #21-24, ARCHITECTURE.md AD-9). Optional
   * and additive — resolved from the originating conversation at preview time
   * (see ChangeSetService.createPreview), so a change set survives its
   * conversation's 24h TTL still knowing which physical file it belongs to.
   * A change set created before this field existed simply has none.
   */
  @Prop({ type: String, required: false, index: true })
  workbookId?: string;

  @Prop({ type: String, required: true, index: true })
  traceId!: string;

  @Prop({ type: Date, required: true, default: () => new Date() })
  timestamp!: Date;

  @Prop({ type: String, required: true })
  prompt!: string;

  @Prop({ type: SchemaTypes.Mixed, default: {} })
  beforeState!: Record<string, CellSnapshotSchema>;

  @Prop({ type: [CellChangeSchema], default: [] })
  changes!: CellChangeSchema[];

  @Prop({ type: [SchemaTypes.Mixed], default: [] })
  actions!: Record<string, unknown>[];

  @Prop({ type: [StructuralOpSchema], default: [] })
  structuralOps!: StructuralOpSchema[];

  /** Distinct action types in this change set with no defined revert path today (TASKS.md #18). */
  @Prop({ type: [String], default: [] })
  irreversibleActionTypes!: string[];

  /** PRD A5 signal — cell changes outside this batch's declared sheet scope (TASKS.md #48). */
  @Prop({ type: [CellChangeSchema], default: [] })
  unintendedChanges!: CellChangeSchema[];

  /** PRD A6 signal — Excel error strings this change set introduced (TASKS.md #49). */
  @Prop({ type: [FormulaErrorChangeSchema], default: [] })
  formulaErrorsIntroduced!: FormulaErrorChangeSchema[];

  @Prop({
    type: String,
    required: true,
    enum: ['previewed', 'applied', 'reverted'],
    default: 'previewed',
  })
  status!: string;

  @Prop({ type: Date })
  appliedAt?: Date;

  @Prop({ type: Date })
  revertedAt?: Date;

  @Prop({ type: Number, required: false })
  provenanceConfidence?: number;

  /**
   * TASKS.md #99 — true once any part of `changes`/`beforeState` came from a
   * frontend-reported real Excel read (`markApplied`'s `frontendChanges`
   * param, e.g. SORT_RANGE on a non-sparse range) rather than the backend's
   * own shadow-workbook diff. Revert's forward-replay self-verification
   * assumes `virtualApply` can accurately re-simulate every action type — a
   * false assumption for exactly the actions that needed a frontend-reported
   * fallback in the first place — so that check is skipped for change sets
   * carrying this flag; the frontend-reported data is ground truth already,
   * re-deriving and comparing against a simulation of it adds no real safety
   * and produces false "would not converge" refusals.
   */
  @Prop({ type: Boolean, default: false })
  hasFrontendReportedChanges?: boolean;
}

export const ChangeSetSchema = SchemaFactory.createForClass(ChangeSet);
