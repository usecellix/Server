import mongoose from 'mongoose';
import { AdminPromptsService } from '../src/admin/admin-prompts.service';

const MONGO_URL = process.env.MONGODB_URL ?? 'mongodb://127.0.0.1:27017';
const TEST_DB = 'cellix_test_admin_prompts';

describe('AdminPromptsService (live Mongo)', () => {
  let connection: mongoose.Connection | null = null;
  let service: AdminPromptsService;
  let userId: string;

  beforeAll(async () => {
    try {
      connection = await mongoose.createConnection(`${MONGO_URL}/${TEST_DB}`, { serverSelectionTimeoutMS: 1500 }).asPromise();
    } catch {
      connection = null;
      return;
    }

    service = new AdminPromptsService(connection, { creditsPerUsd: 600 } as never);

    const userDoc = await connection.db!.collection('user').insertOne({ name: 'Prompt Tester', email: 'admin-prompts-spec@example.com' });
    userId = userDoc.insertedId.toHexString();

    await connection.db!.collection('ai_prompts').insertOne({
      promptId: 'prompt-spec-1',
      prompt: 'Summarize this sheet',
      userId,
      conversationId: 'conv-1',
      createdAt: new Date('2026-01-01'),
      llmCalls: 2,
      totalTokens: 1000,
      costUsd: 0.01,
      lastOutcome: 'ok',
    });
    await connection.db!.collection('llm_calls').insertMany([
      {
        promptId: 'prompt-spec-1',
        ts: new Date('2026-01-01T00:00:01Z'),
        model: 'gpt-5',
        caller: 'planner',
        attempt: 1,
        promptTokens: 400,
        completionTokens: 100,
        totalTokens: 500,
        costUsd: 0.006,
        latencyMs: 800,
        success: true,
      },
      {
        promptId: 'prompt-spec-1',
        ts: new Date('2026-01-01T00:00:02Z'),
        model: 'gpt-5',
        caller: 'executor',
        attempt: 1,
        promptTokens: 400,
        completionTokens: 100,
        totalTokens: 500,
        costUsd: 0.004,
        latencyMs: 500,
        success: true,
      },
    ]);
    await connection.db!.collection('credit_ledger').insertMany([
      { promptId: 'prompt-spec-1', billingEntityId: userId, entryType: 'debit', amount: -6, bucket: 'planCredits', createdAt: new Date() },
    ]);
  }, 20000);

  afterAll(async () => {
    if (connection) {
      await connection.dropDatabase();
      await connection.close();
    }
  });

  it('listPrompts includes creditsCharged per row and a credits sum across the page', async () => {
    if (!connection) {
      console.warn('No MongoDB reachable — skipping live AdminPromptsService check');
      return;
    }
    const { rows, sums } = await service.listPrompts({ sort: 'recent', page: 1 });
    const row = rows.find((r) => r.promptId === 'prompt-spec-1');
    expect(row).toEqual(expect.objectContaining({ creditsCharged: 6, llmCalls: 2 }));
    expect(sums.credits).toBeGreaterThanOrEqual(6);
  }, 20000);

  it('getPrompt returns per-call breakdown, creditsEquivalent, and the same creditsCharged', async () => {
    if (!connection) return;
    const detail = await service.getPrompt('prompt-spec-1');
    expect(detail).not.toBeNull();
    expect(detail!.prompt.creditsCharged).toBe(6);
    expect(detail!.calls).toHaveLength(2);
    expect(detail!.calls[0].creditsEquivalent).toBeCloseTo(0.006 * 600, 5);
    expect(detail!.byCaller.map((r) => r.key).sort()).toEqual(['executor', 'planner']);
  }, 20000);

  it('getPrompt returns null for an unknown id', async () => {
    if (!connection) return;
    await expect(service.getPrompt('does-not-exist')).resolves.toBeNull();
  }, 20000);
});
