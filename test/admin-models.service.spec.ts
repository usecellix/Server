import mongoose from 'mongoose';
import { AdminModelsService } from '../src/admin/admin-models.service';

const MONGO_URL = process.env.MONGODB_URL ?? 'mongodb://127.0.0.1:27017';
const TEST_DB = 'cellix_test_admin_models';

describe('AdminModelsService (live Mongo)', () => {
  let connection: mongoose.Connection | null = null;
  let service: AdminModelsService;

  beforeAll(async () => {
    try {
      connection = await mongoose.createConnection(`${MONGO_URL}/${TEST_DB}`, { serverSelectionTimeoutMS: 1500 }).asPromise();
    } catch {
      connection = null;
      return;
    }
    service = new AdminModelsService(connection);

    await connection.db!.collection('llm_calls').insertMany([
      { model: 'gpt-5', caller: 'planner', ts: new Date(), costUsd: 0.02, totalTokens: 400, success: true, attempt: 1, latencyMs: 900 },
      { model: 'gpt-5', caller: 'executor', ts: new Date(), costUsd: 0.01, totalTokens: 200, success: false, attempt: 2, latencyMs: 1200 },
      { model: 'gpt-5-mini', caller: 'router', ts: new Date(), costUsd: 0.001, totalTokens: 50, success: true, attempt: 1, latencyMs: 100 },
    ]);
  }, 20000);

  afterAll(async () => {
    if (connection) {
      await connection.dropDatabase();
      await connection.close();
    }
  });

  it('groups usage by model and by caller, and keeps a failed/retried call in the totals', async () => {
    if (!connection) {
      console.warn('No MongoDB reachable — skipping live AdminModelsService check');
      return;
    }
    const usage = await service.getModelUsage(null);
    expect(usage).not.toBeNull();
    const gpt5 = usage!.byModel.find((r) => r.key === 'gpt-5');
    expect(gpt5).toEqual(expect.objectContaining({ calls: 2, failed: 1 }));
    expect(usage!.byCaller.map((r) => r.key).sort()).toEqual(['executor', 'planner', 'router']);
    expect(usage!.totals?.calls).toBeGreaterThanOrEqual(3);
  }, 20000);

  it('filters by key (q) against the grouped rows, not the raw calls', async () => {
    if (!connection) return;
    const usage = await service.getModelUsage(null, { q: 'mini' });
    expect(usage!.byModel.map((r) => r.key)).toEqual(['gpt-5-mini']);
  }, 20000);
});
