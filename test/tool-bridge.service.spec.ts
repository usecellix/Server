import { ToolBridgeService } from '../src/agents/tool-bridge.service';

describe('ToolBridgeService', () => {
  it('resolves when deliverResult is called with matching request id', async () => {
    const bridge = new ToolBridgeService();
    let capturedRequestId = '';

    const pending = bridge.waitForRangeData(
      'conv_1',
      { name: 'get_range_data', sheet: 'Sheet1', range: 'A1:B2' },
      (_event, data) => {
        capturedRequestId = String(data.requestId);
      },
    );

    await Promise.resolve();
    expect(capturedRequestId).toMatch(/^tr_/);

    const accepted = bridge.deliverResult('conv_1', capturedRequestId, {
      values: [['A', 'B']],
    });

    expect(accepted).toBe(true);
    await expect(pending).resolves.toEqual({ values: [['A', 'B']] });
  });
});
