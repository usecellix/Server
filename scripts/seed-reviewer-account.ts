/**
 * Creates (or tops up) the AppSource/Partner Center reviewer account:
 * reviewer@usecellix.com, signed in via email+password, pre-loaded with a
 * large credit balance so a Microsoft reviewer can exercise every feature
 * without hitting a paywall mid-review.
 *
 * Does NOT create the Better Auth user itself — that must go through the
 * real signup flow (POST /api/auth/sign-up/email) so the password is
 * hashed correctly by Better Auth, not guessed at here. Run that first:
 *
 *   curl -X POST https://www.usecellix.com/api/auth/sign-up/email \
 *     -H "Content-Type: application/json" \
 *     -d '{"email":"reviewer@usecellix.com","password":"<password>","name":"Cellix Reviewer"}'
 *
 * That response includes the new user's `id` — pass it as REVIEWER_USER_ID
 * below, or let this script look it up by email from the `user` collection
 * itself (default; no id needed as long as the account already exists).
 *
 * Grants through the same shapes CreditLedgerService.grantPlanCredits /
 * addPurchasedCredits produce (credit_accounts + a matching credit_ledger
 * grant row) — never a bare balance edit with no ledger trail. Safe to
 * re-run: it resets planCredits to REVIEWER_CREDITS (same "grant resets,
 * does not stack" rule the real subscription-renewal grant uses, TASKS.md
 * #342) and records the delta as its own ledger row each time, so re-running
 * never silently doubles the balance.
 *
 * Usage (from cellix_backend/):
 *   npx ts-node scripts/seed-reviewer-account.ts
 *   REVIEWER_EMAIL=someone@else.com REVIEWER_CREDITS=5000 npx ts-node scripts/seed-reviewer-account.ts
 *
 * Reads MONGODB_URL / MONGODB_DB_NAME from cellix_backend/.env — point that
 * file at production before running this against production data.
 */
import * as path from 'path';
import * as dotenv from 'dotenv';
import mongoose from 'mongoose';

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

const REVIEWER_EMAIL = process.env.REVIEWER_EMAIL || 'reviewer@usecellix.com';
const REVIEWER_CREDITS = Number(process.env.REVIEWER_CREDITS || 10000);
const REVIEWER_PLAN_TIER = process.env.REVIEWER_PLAN_TIER || 'beta';

const mongoUrl = process.env.MONGODB_URL || 'mongodb://127.0.0.1:27017/cellix';
const dbName = process.env.MONGODB_DB_NAME || 'cellix';

async function main() {
  if (!Number.isFinite(REVIEWER_CREDITS) || REVIEWER_CREDITS <= 0) {
    console.error(`REVIEWER_CREDITS must be a positive number, got: ${process.env.REVIEWER_CREDITS}`);
    process.exit(1);
  }

  console.log(`Connecting to ${mongoUrl.replace(/\/\/[^@]+@/, '//***@')} (db: ${dbName})...`);
  const connection = await mongoose.createConnection(mongoUrl, { dbName }).asPromise();
  const db = connection.db;
  if (!db) throw new Error('No database handle after connecting');

  try {
    const user = await db.collection('user').findOne({ email: REVIEWER_EMAIL });
    if (!user) {
      console.error(
        `No user found with email ${REVIEWER_EMAIL}. Create it first via the real signup flow — see this ` +
          `script's docblock — then re-run.`,
      );
      process.exit(1);
    }
    const billingEntityId = String(user._id);
    console.log(`Found user ${REVIEWER_EMAIL} -> billingEntityId ${billingEntityId}`);

    const existing = await db.collection('credit_accounts').findOne({ billingEntityId });
    const before = existing
      ? Number(existing.planCredits || 0) + Number(existing.purchasedCredits || 0) + Number(existing.oneTimeCredits || 0)
      : 0;

    await db.collection('credit_accounts').updateOne(
      { billingEntityId },
      {
        $set: {
          billingEntityType: 'user',
          billingEntityId,
          planTier: REVIEWER_PLAN_TIER,
          planCredits: REVIEWER_CREDITS,
          purchasedCredits: 0,
          oneTimeCredits: 0,
        },
      },
      { upsert: true },
    );

    await db.collection('credit_ledger').insertOne({
      billingEntityId,
      entryType: 'grant',
      amount: REVIEWER_CREDITS,
      bucket: 'planCredits',
      createdAt: new Date(),
    });
    if (before > 0) {
      // Same "reset, log the forfeited amount as its own row" rule
      // grantPlanCredits uses — see credit-ledger.service.ts.
      await db.collection('credit_ledger').insertOne({
        billingEntityId,
        entryType: 'expire',
        amount: -before,
        bucket: 'planCredits',
        createdAt: new Date(),
      });
    }

    console.log(
      `Set ${REVIEWER_EMAIL}'s balance to ${REVIEWER_CREDITS} credits (plan tier: ${REVIEWER_PLAN_TIER}).` +
        (before > 0 ? ` Previous balance (${before}) logged as expired, not stacked.` : ''),
    );
  } finally {
    await connection.close();
  }
}

main().catch((error) => {
  console.error('Failed to seed reviewer account:', error);
  process.exit(1);
});
