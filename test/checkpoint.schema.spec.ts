import mongoose from 'mongoose';
import { Checkpoint, CheckpointSchema } from '../src/audit/schemas/checkpoint.schema';

describe('Checkpoint schema (TASKS.md #26)', () => {
  const CheckpointModel = mongoose.model<Checkpoint>(
    `CheckpointTest_${Date.now()}`,
    CheckpointSchema,
  );

  it('a well-formed checkpoint document round-trips', () => {
    const doc = new CheckpointModel({
      checkpointId: 'cp-1',
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'before risky delete',
      trigger: 'auto',
      anchorChangeSetId: 'cs-1',
      createdAt: new Date(),
      status: 'active',
    });

    expect(doc.validateSync()).toBeUndefined();
    const obj = doc.toObject();
    expect(obj).toMatchObject({
      checkpointId: 'cp-1',
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'before risky delete',
      trigger: 'auto',
      anchorChangeSetId: 'cs-1',
      status: 'active',
    });
  });

  it('accepts an empty anchorChangeSetId (meaning "beginning of history")', () => {
    const doc = new CheckpointModel({
      checkpointId: 'cp-2',
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'first checkpoint',
      trigger: 'manual',
      anchorChangeSetId: '',
      status: 'active',
    });
    expect(doc.validateSync()).toBeUndefined();
  });

  it('a document missing workbookId is rejected', () => {
    const doc = new CheckpointModel({
      checkpointId: 'cp-3',
      conversationId: 'conv-1',
      label: 'missing workbookId',
      trigger: 'manual',
      anchorChangeSetId: 'cs-1',
      status: 'active',
    });
    expect(doc.validateSync()?.errors.workbookId).toBeDefined();
  });

  it('a document with no anchorChangeSetId defaults to "" (beginning of history), not rejected', () => {
    // '' is a real, meaningful value here (see the schema's own comment) — Mongoose's
    // String `required: true` rejects empty string, which would make "no prior applied
    // change sets" unrepresentable, so this field is `required: false, default: ''`.
    const doc = new CheckpointModel({
      checkpointId: 'cp-4',
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'missing anchor',
      trigger: 'manual',
      status: 'active',
    });
    expect(doc.validateSync()).toBeUndefined();
    expect(doc.anchorChangeSetId).toBe('');
  });

  it('a document missing checkpointId is rejected', () => {
    const doc = new CheckpointModel({
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'missing id',
      trigger: 'manual',
      anchorChangeSetId: 'cs-1',
      status: 'active',
    });
    expect(doc.validateSync()?.errors.checkpointId).toBeDefined();
  });

  it('defaults status to "active" when omitted', () => {
    const doc = new CheckpointModel({
      checkpointId: 'cp-5',
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'default status',
      trigger: 'manual',
      anchorChangeSetId: 'cs-1',
    });
    expect(doc.status).toBe('active');
  });

  it('rejects an unknown trigger value', () => {
    const doc = new CheckpointModel({
      checkpointId: 'cp-6',
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'bad trigger',
      trigger: 'scheduled',
      anchorChangeSetId: 'cs-1',
      status: 'active',
    });
    expect(doc.validateSync()?.errors.trigger).toBeDefined();
  });
});
