import mongoose from 'mongoose';
import { AdminOverviewService } from '../src/admin/admin-overview.service';

const MONGO_URL = process.env.MONGODB_URL ?? 'mongodb://127.0.0.1:27017';
const TEST_DB = 'cellix_test_admin_overview';

describe('AdminOverviewService (live Mongo)', () => {
  let connection: mongoose.Connection | null = null;
  let service: AdminOverviewService;

  beforeAll(async () => {
    try {
      connection = await mongoose.createConnection(`${MONGO_URL}/${TEST_DB}`, { serverSelectionTimeoutMS: 1500 }).asPromise();
    } catch {
      connection = null;
      return;
    }
    service = new AdminOverviewService(connection);

    const userDoc = await connection.db!.collection('user').insertOne({ name: 'Overview Tester', email: 'admin-overview-spec@example.com', createdAt: new Date() });
    const userId = userDoc.insertedId.toHexString();

    await connection.db!.collection('ai_prompts').insertMany([
      { promptId: 'ov-1', userId, prompt: 'a', createdAt: new Date(), llmCalls: 1, totalTokens: 100, costUsd: 0.01, lastOutcome: 'ok' },
      { promptId: 'ov-2', userId, prompt: 'b', createdAt: new Date(), llmCalls: 1, totalTokens: 100, costUsd: 0.02, lastOutcome: 'error' },
    ]);
    await connection.db!.collection('llm_calls').insertOne({ model: 'gpt-5', ts: new Date(), costUsd: 0.03, totalTokens: 200 });
  }, 20000);

  afterAll(async () => {
    if (connection) {
      await connection.dropDatabase();
      await connection.close();
    }
  });

  it('getOverview totals prompts/cost/errors and lists recent prompts with a spend series', async () => {
    if (!connection) {
      console.warn('No MongoDB reachable — skipping live AdminOverviewService check');
      return;
    }
    const overview = await service.getOverview({ from: null, to: new Date(), bucket: 'day' });
    expect(overview).not.toBeNull();
    expect(overview!.totals.prompts).toBeGreaterThanOrEqual(2);
    expect(overview!.totals.erroredPrompts).toBeGreaterThanOrEqual(1);
    expect(overview!.recentPrompts.length).toBeGreaterThanOrEqual(2);
    expect(overview!.spendSeries.length).toBeGreaterThan(0);
  }, 20000);
});
