import mongoose from 'mongoose';

/**
 * TASKS.md #233 — the credit debit is an aggregation-PIPELINE update (an array),
 * and Mongoose 9 rejects one unless `updatePipeline: true` is passed. Every
 * billable request from a signed-in user threw "Cannot pass an array to query
 * updates unless the `updatePipeline` option is set" at debit time, after the
 * plan had already been produced and verified — the task pane showed a red
 * error instead of the change.
 *
 * `credit-ledger.service.spec.ts` could not catch it: it mocks
 * findOneAndUpdate, so the option never reaches a driver that cares. This test
 * runs the real call against the local MongoDB and is skipped when none is
 * reachable, so CI without Mongo stays green rather than silently vacuous.
 */
const MONGO_URL = process.env.MONGODB_URL ?? 'mongodb://127.0.0.1:27017';
const TEST_DB = 'cellix_test_pipeline_update';

describe('credit debit pipeline update against a real MongoDB (#229)', () => {
  let connection: mongoose.Connection | null = null;

  beforeAll(async () => {
    try {
      connection = await mongoose
        .createConnection(`${MONGO_URL}/${TEST_DB}`, { serverSelectionTimeoutMS: 1500 })
        .asPromise();
    } catch {
      connection = null;
    }
  }, 20000);

  afterAll(async () => {
    if (connection) {
      await connection.dropDatabase();
      await connection.close();
    }
  });

  it('applies a $set pipeline update when updatePipeline is set', async () => {
    if (!connection) {
      console.warn('No MongoDB reachable — skipping live pipeline-update check');
      return;
    }

    const model = connection.model(
      'PipelineProbe',
      new mongoose.Schema({ billingEntityId: String, planCredits: Number, oneTimeCredits: Number }),
    );
    await model.create({ billingEntityId: 'probe', planCredits: 10, oneTimeCredits: 5 });

    const cost = 4;
    const updated = await model.findOneAndUpdate(
      {
        billingEntityId: 'probe',
        $expr: { $gte: [{ $add: ['$planCredits', '$oneTimeCredits'] }, cost] },
      },
      [{ $set: { planCredits: { $max: [0, { $subtract: ['$planCredits', cost] }] } } }],
      { new: true, updatePipeline: true },
    );

    expect(updated).not.toBeNull();
    expect(updated!.get('planCredits')).toBe(6);
  }, 20000);

  it('is exactly the call that fails without the flag', async () => {
    if (!connection) return;

    const model = connection.model('PipelineProbe');
    // Mongoose raises this SYNCHRONOUSLY from the call itself, not as a
    // rejected query — which is why the production failure surfaced as a
    // thrown error mid-request rather than a failed await.
    let raised: unknown;
    try {
      await model.findOneAndUpdate({ billingEntityId: 'probe' }, [{ $set: { planCredits: 1 } }], {
        new: true,
      } as never);
    } catch (error) {
      raised = error;
    }
    expect(String((raised as Error)?.message)).toMatch(/updatePipeline/i);
  }, 20000);
});
