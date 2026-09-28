import mongoose from 'mongoose';
import { AdminBillingService } from '../src/admin/admin-billing.service';

const MONGO_URL = process.env.MONGODB_URL ?? 'mongodb://127.0.0.1:27017';
const TEST_DB = 'cellix_test_admin_billing';

describe('AdminBillingService (live Mongo)', () => {
  let connection: mongoose.Connection | null = null;
  let service: AdminBillingService;
  let userId: string;

  beforeAll(async () => {
    try {
      connection = await mongoose.createConnection(`${MONGO_URL}/${TEST_DB}`, { serverSelectionTimeoutMS: 1500 }).asPromise();
    } catch {
      connection = null;
      return;
    }
    service = new AdminBillingService(connection);

    const userDoc = await connection.db!.collection('user').insertOne({ name: 'Billing Tester', email: 'admin-billing-spec@example.com' });
    userId = userDoc.insertedId.toHexString();

    await connection.db!.collection('subscriptions').insertOne({
      billingEntityId: userId,
      razorpaySubscriptionId: 'sub_spec_1',
      status: 'active',
      planTier: 'solo',
      currentPeriodEnd: new Date('2026-12-01'),
      cancelAtPeriodEnd: false,
      createdAt: new Date('2026-01-01'),
    });
    await connection.db!.collection('credit_accounts').insertOne({
      billingEntityType: 'user',
      billingEntityId: userId,
      planTier: 'solo',
      planCredits: 3000,
      purchasedCredits: 0,
      oneTimeCredits: 0,
    });
    // A guest (email-keyed) account, to exercise the owner/guest classification.
    await connection.db!.collection('credit_accounts').insertOne({
      billingEntityType: 'user',
      billingEntityId: 'guest-spec@example.com',
      planTier: 'free',
      planCredits: 0,
      purchasedCredits: 0,
      oneTimeCredits: 120,
    });
    await connection.db!.collection('credit_ledger').insertOne({
      billingEntityId: userId,
      entryType: 'debit',
      amount: -10,
      bucket: 'planCredits',
      createdAt: new Date(),
    });
  }, 20000);

  afterAll(async () => {
    if (connection) {
      await connection.dropDatabase();
      await connection.close();
    }
  });

  it('getBillingSummary counts the active solo subscription and both credit accounts', async () => {
    if (!connection) {
      console.warn('No MongoDB reachable — skipping live AdminBillingService check');
      return;
    }
    const summary = await service.getBillingSummary(null);
    expect(summary).not.toBeNull();
    expect(summary!.activeCount).toBeGreaterThanOrEqual(1);
    expect(summary!.activeByPlan.solo).toBeGreaterThanOrEqual(1);
    expect(summary!.accounts.total).toBeGreaterThanOrEqual(2);
    expect(summary!.accounts.guest).toBeGreaterThanOrEqual(1);
  }, 20000);

  it('listSubscriptions resolves the owner to the real signed-in user, not just the raw id', async () => {
    if (!connection) return;
    const { rows } = await service.listSubscriptions({ since: null, sort: 'recent', page: 1 });
    const row = rows.find((r) => r.billingEntityId === userId);
    expect(row?.owner).toEqual(expect.objectContaining({ guest: false, userId, email: 'admin-billing-spec@example.com' }));
  }, 20000);

  it('listLedger returns the debit row with its owner attached', async () => {
    if (!connection) return;
    const { rows, total } = await service.listLedger({ since: null, sort: 'recent', page: 1 });
    expect(total).toBeGreaterThanOrEqual(1);
    const row = rows.find((r) => r.billingEntityId === userId);
    expect(row).toEqual(expect.objectContaining({ amount: -10, owner: expect.objectContaining({ guest: false }) }));
  }, 20000);
});
