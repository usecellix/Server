import mongoose from 'mongoose';
import { ChangeSet, ChangeSetSchema } from '../src/audit/schemas/change-set.schema';

describe('ChangeSet schema — structuralOps (TASKS.md #11)', () => {
  const ChangeSetModel = mongoose.model<ChangeSet>(
    `ChangeSetTest_${Date.now()}`,
    ChangeSetSchema,
  );

  it('defaults structuralOps to [] for a document that never had the field (pre-migration data)', () => {
    // Simulate a document loaded straight from Mongo before this field existed —
    // no structuralOps key present in the raw source at all.
    const legacyRaw = {
      changeSetId: 'cs-legacy',
      conversationId: 'conv-1',
      traceId: 'trace-1',
      timestamp: new Date(),
      prompt: 'sort by total',
      beforeState: {},
      changes: [],
      actions: [],
      status: 'applied',
    };

    const doc = new ChangeSetModel(legacyRaw);
    expect(doc.structuralOps).toEqual([]);
    expect(doc.validateSync()).toBeUndefined();
  });

  it('accepts and round-trips a populated structuralOps array', () => {
    const doc = new ChangeSetModel({
      changeSetId: 'cs-new',
      conversationId: 'conv-2',
      traceId: 'trace-2',
      timestamp: new Date(),
      prompt: 'add a new sheet',
      beforeState: {},
      changes: [],
      actions: [],
      structuralOps: [
        {
          opType: 'ADD_SHEET',
          sheetName: 'January',
          params: { position: 3 },
          appliedAt: new Date(),
        },
      ],
      status: 'applied',
    });

    expect(doc.validateSync()).toBeUndefined();
    const obj = doc.toObject();
    expect(obj.structuralOps).toHaveLength(1);
    expect(obj.structuralOps[0]).toMatchObject({
      opType: 'ADD_SHEET',
      sheetName: 'January',
      params: { position: 3 },
    });
  });
});

describe('ChangeSet schema — irreversibleActionTypes (TASKS.md #18)', () => {
  const ChangeSetModel = mongoose.model<ChangeSet>(
    `ChangeSetTest_${Date.now()}_irreversible`,
    ChangeSetSchema,
  );

  it('defaults irreversibleActionTypes to [] for a document that never had the field', () => {
    const doc = new ChangeSetModel({
      changeSetId: 'cs-legacy-2',
      conversationId: 'conv-1',
      traceId: 'trace-1',
      timestamp: new Date(),
      prompt: 'freeze panes',
      beforeState: {},
      changes: [],
      actions: [],
      status: 'applied',
    });
    expect(doc.irreversibleActionTypes).toEqual([]);
    expect(doc.validateSync()).toBeUndefined();
  });

  it('accepts and round-trips a populated irreversibleActionTypes array', () => {
    const doc = new ChangeSetModel({
      changeSetId: 'cs-new-2',
      conversationId: 'conv-2',
      traceId: 'trace-2',
      timestamp: new Date(),
      prompt: 'freeze panes and rename sheet',
      beforeState: {},
      changes: [],
      actions: [],
      irreversibleActionTypes: ['FREEZE_PANES', 'RENAME_SHEET'],
      status: 'applied',
    });

    expect(doc.validateSync()).toBeUndefined();
    expect(doc.toObject().irreversibleActionTypes).toEqual(['FREEZE_PANES', 'RENAME_SHEET']);
  });
});
