import { Injectable, Logger } from '@nestjs/common';
import { InjectConnection, InjectModel } from '@nestjs/mongoose';
import { Connection, Model } from 'mongoose';
import { CreditAccount, CreditAccountDocument } from './schemas/credit-account.schema';
import { CreditLedgerEntry, CreditLedgerEntryDocument } from './schemas/credit-ledger.schema';
import { Subscription, SubscriptionDocument } from './schemas/subscription.schema';

/**
 * The marketing site's guest checkout keys a credit account by the payer's
 * email (RazorpayCheckoutService.createGuestSubscriptionSession), while the
 * signed-in product keys it by user id. Without a link between the two, a
 * customer who paid before signing in never saw the credits they bought, and
 * every renewal kept granting onto the orphaned email account.
 */
@Injectable()
export class GuestAccountLinkService {
  private readonly logger = new Logger(GuestAccountLinkService.name);

  constructor(
    @InjectConnection() private readonly connection: Connection,
    @InjectModel(CreditAccount.name)
    private readonly creditAccountModel: Model<CreditAccountDocument>,
    @InjectModel(CreditLedgerEntry.name)
    private readonly creditLedgerModel: Model<CreditLedgerEntryDocument>,
    @InjectModel(Subscription.name)
    private readonly subscriptionModel: Model<SubscriptionDocument>,
  ) {}

  /**
   * Maps an email-shaped billingEntityId (from Razorpay notes) to the
   * signed-in user with that email, so webhook grants land on the account the
   * product actually reads. Anything else passes through unchanged.
   */
  async resolveBillingEntityId(billingEntityId: string): Promise<string> {
    if (!billingEntityId.includes('@')) return billingEntityId;
    const user = await this.connection
      .collection('user')
      .findOne({ email: billingEntityId.trim().toLowerCase() }, { projection: { _id: 1 } });
    return user ? String(user._id) : billingEntityId;
  }

  /**
   * Moves whatever a guest checkout left on the email-keyed account onto the
   * signed-in user's account: plan tier, plan credits, purchased credits, and
   * the subscription rows. The guest's own Free-tier one-time grant stays
   * behind — the user already received theirs. Claimed atomically (the guest
   * account is zeroed and stamped in the same update), so concurrent balance
   * reads can't move it twice. Returns the credits moved.
   */
  async claimGuestAccount(userId: string, email: string | undefined | null): Promise<number> {
    const guestId = email?.trim().toLowerCase();
    if (!guestId || guestId === userId) return 0;

    const guest = await this.creditAccountModel
      .findOneAndUpdate(
        {
          billingEntityId: guestId,
          linkedToUserId: { $exists: false },
          $or: [{ planCredits: { $gt: 0 } }, { purchasedCredits: { $gt: 0 } }, { planTier: { $ne: 'free' } }],
        },
        { $set: { planCredits: 0, purchasedCredits: 0, planTier: 'free', linkedToUserId: userId } },
        { returnDocument: 'before', projection: { planCredits: 1, purchasedCredits: 1, planTier: 1 } },
      )
      .lean();
    if (!guest) return 0;

    const planCredits = Math.max(0, guest.planCredits ?? 0);
    const purchasedCredits = Math.max(0, guest.purchasedCredits ?? 0);
    await this.creditAccountModel.updateOne(
      { billingEntityId: userId },
      {
        $inc: { planCredits, purchasedCredits },
        ...(guest.planTier && guest.planTier !== 'free' ? { $set: { planTier: guest.planTier } } : {}),
      },
    );
    await this.subscriptionModel.updateMany({ billingEntityId: guestId }, { $set: { billingEntityId: userId } });

    const now = new Date();
    const rows: Record<string, unknown>[] = [];
    for (const [bucket, amount] of [['planCredits', planCredits], ['purchasedCredits', purchasedCredits]] as const) {
      if (amount <= 0) continue;
      rows.push(
        { billingEntityId: userId, entryType: bucket === 'planCredits' ? 'grant' : 'purchase', amount, bucket, createdAt: now },
        { billingEntityId: guestId, entryType: 'expire', amount: -amount, bucket, createdAt: now },
      );
    }
    if (rows.length) await this.creditLedgerModel.insertMany(rows);

    this.logger.log(
      `Linked guest account ${guestId} to user ${userId}: ${planCredits} plan + ${purchasedCredits} purchased credits, tier ${guest.planTier}`,
    );
    return planCredits + purchasedCredits;
  }
}
