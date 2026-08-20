import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';

/**
 * TASKS.md #25 — WorkflowTrace.workbookId. startTrace() is fire-and-forget
 * (public method swallows its own promise), so these tests await the async
 * work via the mocked model call directly rather than the public method's
 * return value.
 */
function buildService() {
  const findOneAndUpdate = jest.fn().mockResolvedValue(undefined);
  const model = { findOneAndUpdate };
  const service = new WorkflowTraceService(model as never);
  return { service, findOneAndUpdate };
}

describe('WorkflowTraceService.startTrace — workbookId (TASKS.md #25)', () => {
  it('includes workbookId in the inserted document when provided', async () => {
    const { service, findOneAndUpdate } = buildService();

    await (service as unknown as { startTraceAsync: (input: unknown) => Promise<void> })
      .startTraceAsync({
        traceId: 'trace-1',
        conversationId: 'conv-1',
        workbookId: 'wb_shared-1',
        message: 'add a total column',
      });

    expect(findOneAndUpdate).toHaveBeenCalledWith(
      { traceId: 'trace-1' },
      expect.objectContaining({
        $setOnInsert: expect.objectContaining({ workbookId: 'wb_shared-1' }),
      }),
      { upsert: true },
    );
  });

  it('omits workbookId entirely when not provided (backward compatible)', async () => {
    const { service, findOneAndUpdate } = buildService();

    await (service as unknown as { startTraceAsync: (input: unknown) => Promise<void> })
      .startTraceAsync({
        traceId: 'trace-2',
        conversationId: 'conv-2',
        message: 'add a total column',
      });

    const call = findOneAndUpdate.mock.calls[0];
    const setOnInsert = (call[1] as { $setOnInsert: Record<string, unknown> }).$setOnInsert;
    expect(setOnInsert).not.toHaveProperty('workbookId');
  });
});
