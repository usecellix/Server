import { AuditService } from '../src/audit/audit.service';

describe('AuditService export helpers', () => {
  it('builds CSV export with change sets and llm calls', () => {
    const service = Object.create(AuditService.prototype) as AuditService;
    const fromDate = new Date('2024-01-01T00:00:00Z');
    const toDate = new Date('2024-01-02T00:00:00Z');
    const csv = service.buildExportCsv(fromDate, toDate, {
      exportedAt: '2024-01-02T00:00:00Z',
      from: fromDate.toISOString(),
      to: toDate.toISOString(),
      changeSets: [
        {
          changeSetId: 'cs_1',
          conversationId: 'conv_1',
          traceId: 'trace_1',
          timestamp: fromDate,
          prompt: 'Add row',
          beforeState: {},
          changes: [],
          actions: [],
          status: 'applied',
        },
      ],
      auditLogs: [
        {
          id: 'log_1',
          traceId: 'trace_1',
          timestamp: toDate.toISOString(),
          model: 'gpt-test',
          tier: 'high',
          intent: 'modify_data',
          promptTokens: 100,
          completionTokens: 50,
          totalTokens: 150,
          estimatedCostUsd: 0.002,
          latencyMs: 900,
          success: true,
        },
      ],
    });

    expect(csv).toContain('recordType,id,traceId');
    expect(csv).toContain('change_set,cs_1');
    expect(csv).toContain('llm_call,log_1');
  });
});
