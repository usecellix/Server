import mongoose from 'mongoose';
import { AdminUsersService } from '../src/admin/admin-users.service';
import { CreditAccountSchema } from '../src/credit/schemas/credit-account.schema';
import { CreditLedgerEntrySchema } from '../src/credit/schemas/credit-ledger.schema';
import { SubscriptionSchema } from '../src/credit/schemas/subscription.schema';

/**
 * AdminUsersService reads two shapes: Mongoose models for credit_accounts/
 * credit_ledger/subscriptions, and the `user`/`session`/`account`/
 * `ai_prompts`/`conversations` collections directly off the raw connection
 * (better-auth's own collections have no Mongoose schema in this codebase).
 * Mocking that mix would mostly test the mock, so — like
 * credit-debit-live-mongo.spec.ts — this runs against a real local MongoDB
 * and is skipped when none is reachable.
 */
const MONGO_URL = process.env.MONGODB_URL ?? 'mongodb://127.0.0.1:27017';
const TEST_DB = 'cellix_test_admin_users';

describe('AdminUsersService (live Mongo)', () => {
  let connection: mongoose.Connection | null = null;
  let service: AdminUsersService;
  let userId: string;

  beforeAll(async () => {
    try {
      connection = await mongoose.createConnection(`${MONGO_URL}/${TEST_DB}`, { serverSelectionTimeoutMS: 1500 }).asPromise();
    } catch {
      connection = null;
      return;
    }

    const creditAccountModel = connection.model('CreditAccount', CreditAccountSchema);
    const creditLedgerModel = connection.model('CreditLedgerEntry', CreditLedgerEntrySchema);
    const subscriptionModel = connection.model('Subscription', SubscriptionSchema);
    service = new AdminUsersService(connection, creditAccountModel as never, creditLedgerModel as never, subscriptionModel as never);

    const userDoc = await connection.db!.collection('user').insertOne({
      name: 'Test User',
      email: 'admin-users-spec@example.com',
      createdAt: new Date('2026-01-01'),
    });
    userId = userDoc.insertedId.toHexString();

    await creditAccountModel.create({
      billingEntityType: 'user',
      billingEntityId: userId,
      planTier: 'solo',
      planCredits: 900,
      purchasedCredits: 100,
      oneTimeCredits: 0,
    });
    await creditLedgerModel.create([
      { billingEntityId: userId, entryType: 'debit', amount: -30, bucket: 'planCredits', createdAt: new Date() },
      { billingEntityId: userId, entryType: 'debit', amount: -20, bucket: 'planCredits', createdAt: new Date() },
    ]);
    await connection.db!.collection('ai_prompts').insertOne({
      userId,
      costUsd: 0.05,
      llmCalls: 3,
      totalTokens: 500,
      createdAt: new Date(),
    });
  }, 20000);

  afterAll(async () => {
    if (connection) {
      await connection.dropDatabase();
      await connection.close();
    }
  });

  it('listUsers includes the balance, credits used, and prompt usage for a real user', async () => {
    if (!connection) {
      console.warn('No MongoDB reachable — skipping live AdminUsersService check');
      return;
    }

    const { rows, total } = await service.listUsers({ page: 1, sort: 'recent' });
    expect(total).toBeGreaterThanOrEqual(1);
    const row = rows.find((r) => r.id === userId);
    expect(row).toBeDefined();
    expect(row).toEqual(
      expect.objectContaining({
        email: 'admin-users-spec@example.com',
        plan: 'solo',
        credits: 1000,
        creditsUsed: 50,
        prompts: 1,
      }),
    );
  }, 20000);

  it('getUser returns the same balance/usage figures for one user', async () => {
    if (!connection) return;

    const detail = await service.getUser(userId);
    expect(detail).not.toBeNull();
    expect((detail as any).balance).toEqual(
      expect.objectContaining({ planTier: 'solo', total: 1000 }),
    );
    expect((detail as any).usage.allTime).toEqual(expect.objectContaining({ prompts: 1, creditsUsed: 50 }));
  }, 20000);

  it('getUser returns null for a well-formed id that does not exist', async () => {
    if (!connection) return;
    await expect(service.getUser('6a5f4ed4821d0f40d58d8aff')).resolves.toBeNull();
  }, 20000);

  it('getUser returns null instead of throwing for a malformed id', async () => {
    if (!connection) return;
    await expect(service.getUser('not-an-object-id')).resolves.toBeNull();
  }, 20000);
});
