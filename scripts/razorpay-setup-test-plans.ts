/**
 * One-time setup for exercising the real Razorpay checkout/webhook flow in
 * TEST MODE — no mocking. Razorpay Subscriptions require a `plan_id` that
 * already exists (razorpay-checkout.service.ts's docblock: "unlike a dynamic
 * Stripe Price, a Razorpay Plan is a one-time setup step ... that must exist
 * before this service can create a subscription against it"). This script
 * creates the Beta/Solo/Firm Plans against Razorpay's TEST API using your
 * test-mode key pair, and prints the resulting plan_ids to paste into `.env`.
 *
 * Usage:
 *   1. Sign up at https://dashboard.razorpay.com (free), switch the
 *      dashboard's mode toggle to "Test Mode" (top-left).
 *   2. Settings -> API Keys -> Generate Test Key. Put the pair in
 *      cellix_backend/.env:
 *        RAZORPAY_KEY_ID=rzp_test_xxxxxxxxxxxxxx
 *        RAZORPAY_KEY_SECRET=xxxxxxxxxxxxxxxxxxxxxxxx
 *   3. Run this script:
 *        npx ts-node scripts/razorpay-setup-test-plans.ts
 *   4. Copy the three printed plan_id values into `.env`:
 *        RAZORPAY_PLAN_ID_BETA=plan_xxxxxxxxxxxx
 *        RAZORPAY_PLAN_ID_SOLO=plan_xxxxxxxxxxxx
 *        RAZORPAY_PLAN_ID_FIRM=plan_xxxxxxxxxxxx
 *
 * Safe to re-run: Razorpay has no "get or create by name" lookup, so running
 * this twice creates duplicate Plans (harmless — just extra clutter in the
 * test dashboard). Only run it again if you want fresh plan_ids, or after
 * switching to a different test-mode key pair.
 *
 * This never touches LIVE mode — it fails loudly if RAZORPAY_KEY_ID doesn't
 * start with `rzp_test_`, so pointing this at production keys by accident
 * isn't possible.
 */
import * as path from 'path';
import * as dotenv from 'dotenv';
import Razorpay from 'razorpay';

dotenv.config({ path: path.resolve(__dirname, '..', '.env') });

const keyId = process.env.RAZORPAY_KEY_ID;
const keySecret = process.env.RAZORPAY_KEY_SECRET;

async function main() {
  if (!keyId || !keySecret) {
    console.error('RAZORPAY_KEY_ID / RAZORPAY_KEY_SECRET not set in cellix_backend/.env — see this file\'s docblock.');
    process.exit(1);
  }
  if (!keyId.startsWith('rzp_test_')) {
    console.error(
      `RAZORPAY_KEY_ID (${keyId}) does not look like a test-mode key (expected rzp_test_...). ` +
        'Refusing to run against what may be a live key.',
    );
    process.exit(1);
  }

  const razorpay = new Razorpay({ key_id: keyId, key_secret: keySecret });

  const planDefs = [
    { envVar: 'RAZORPAY_PLAN_ID_BETA', name: 'Cellix Founding Beta', amountPaise: 89900 },
    { envVar: 'RAZORPAY_PLAN_ID_SOLO', name: 'Cellix CA Professional', amountPaise: 129900 },
    { envVar: 'RAZORPAY_PLAN_ID_FIRM', name: 'Cellix Firm Plan', amountPaise: 599900 },
  ] as const;

  console.log(`Creating ${planDefs.length} test-mode Razorpay Plans (monthly, INR)...\n`);

  const results: string[] = [];
  for (const def of planDefs) {
    const plan = await razorpay.plans.create({
      period: 'monthly',
      interval: 1,
      item: {
        name: def.name,
        amount: def.amountPaise,
        currency: 'INR',
      },
    });
    console.log(`  ${def.name}: ${plan.id}  (₹${def.amountPaise / 100}/mo)`);
    results.push(`${def.envVar}=${plan.id}`);
  }

  console.log('\nPaste these into cellix_backend/.env:\n');
  console.log(results.join('\n'));
}

main().catch((error) => {
  console.error('Failed to create Razorpay test plans:', error);
  process.exit(1);
});
