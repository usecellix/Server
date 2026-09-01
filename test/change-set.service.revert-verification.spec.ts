import { ChangeSetService } from '../src/audit/change-set.service';
import { RevertVerificationError } from '../src/audit/errors/revert-verification.error';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { Model } from 'mongoose';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';

/** Minimal Mongoose-document-shaped fake — enough for revert()'s read/mutate/save pattern. */
function fakeDoc(overrides: Record<string, unknown>) {
  const doc: Record<string, unknown> & { save: jest.Mock } = {
    changeSetId: 'cs-1',
    status: 'applied',
    beforeState: {},
    changes: [],
    actions: [],
    structuralOps: [],
    revertedAt: undefined,
    save: jest.fn().mockResolvedValue(undefined),
    ...overrides,
  };
  return doc;
}

function buildService(doc: ReturnType<typeof fakeDoc>) {
  const model = { findOne: jest.fn().mockResolvedValue(doc) } as unknown as Model<ChangeSetDocument>;
  const workflowTrace = {
    appendTerminalByChangeSet: jest.fn(),
    appendTerminalByConversationId: jest.fn(),
  } as unknown as WorkflowTraceService;
  return new ChangeSetService(model, {} as never, workflowTrace);
}

describe('ChangeSetService.revert — fail-closed self-verification (TASKS.md #19)', () => {
  it('succeeds when the captured changes fully account for what the actions did', async () => {
    const doc = fakeDoc({
      beforeState: {
        'Sheet1!A1': { value: 'old', formula: '', format: 'General' },
      },
      actions: [{ type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 'new' }],
      changes: [{ cell: 'A1', sheet: 'Sheet1', before: 'old', after: 'new', isHardcoded: true }],
    });
    const service = buildService(doc);

    const result = await service.revert('cs-1');

    expect(doc.status).toBe('reverted');
    expect(doc.save).toHaveBeenCalled();
    expect(result.inverseActions.length).toBeGreaterThan(0);
  });

  it('refuses the revert (fail closed) when one cell change is missing from the captured diff — the whole revert is blocked, not partially applied', async () => {
    // Two cells were genuinely changed by `actions`, but `changes` only captured one —
    // simulating a broken/incomplete diff capture, per the task's own acceptance wording.
    const doc = fakeDoc({
      beforeState: {
        'Sheet1!A1': { value: 'old', formula: '', format: 'General' },
        'Sheet1!B1': { value: 'old2', formula: '', format: 'General' },
      },
      actions: [
        { type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 'new' },
        { type: 'SET_CELL', sheetName: 'Sheet1', address: 'B1', value: 'new2' },
      ],
      // B1's change entry is missing — beforeStateToInverseActions will never build an
      // inverse for it, so it stays stuck at 'new2' after the "revert".
      changes: [{ cell: 'A1', sheet: 'Sheet1', before: 'old', after: 'new', isHardcoded: true }],
    });
    const service = buildService(doc);

    await expect(service.revert('cs-1')).rejects.toThrow(RevertVerificationError);

    // Fail closed: no partial success — status untouched, nothing saved.
    expect(doc.status).toBe('applied');
    expect(doc.save).not.toHaveBeenCalled();
  });

  it('names the specific blocking cell in the error', async () => {
    const doc = fakeDoc({
      beforeState: {
        'Sheet1!A1': { value: 'old', formula: '', format: 'General' },
        'Sheet1!B1': { value: 'old2', formula: '', format: 'General' },
      },
      actions: [
        { type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 'new' },
        { type: 'SET_CELL', sheetName: 'Sheet1', address: 'B1', value: 'new2' },
      ],
      changes: [{ cell: 'A1', sheet: 'Sheet1', before: 'old', after: 'new', isHardcoded: true }],
    });
    const service = buildService(doc);

    try {
      await service.revert('cs-1');
      fail('expected revert to throw');
    } catch (error) {
      expect(error).toBeInstanceOf(RevertVerificationError);
      const verificationError = error as RevertVerificationError;
      expect(verificationError.blockingChanges).toHaveLength(1);
      expect(verificationError.blockingChanges[0]).toMatchObject({ sheet: 'Sheet1', cell: 'B1' });
      expect(verificationError.message).toContain('Sheet1!B1');
    }
  });
});
