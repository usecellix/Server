import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { CreditAccount, CreditAccountSchema } from './schemas/credit-account.schema';
import { CreditLedgerEntry, CreditLedgerEntrySchema } from './schemas/credit-ledger.schema';
import {
  ProcessedStripeEvent,
  ProcessedStripeEventSchema,
  Subscription,
  SubscriptionSchema,
} from './schemas/subscription.schema';
import { CreditGateService } from './credit-gate.service';
import { CreditLedgerService } from './credit-ledger.service';
import { CreditAccountQueryService } from './credit-account-query.service';
import { StripeCheckoutService } from './stripe-checkout.service';
import { StripeWebhookService } from './stripe-webhook.service';
import { BillingController, PublicBillingController, StripeWebhookController } from './billing.controller';

@Module({
  imports: [
    AppConfigModule,
    MongooseModule.forFeature([
      { name: CreditAccount.name, schema: CreditAccountSchema },
      { name: CreditLedgerEntry.name, schema: CreditLedgerEntrySchema },
      { name: Subscription.name, schema: SubscriptionSchema },
      { name: ProcessedStripeEvent.name, schema: ProcessedStripeEventSchema },
    ]),
  ],
  providers: [
    CreditGateService,
    CreditLedgerService,
    CreditAccountQueryService,
    StripeCheckoutService,
    StripeWebhookService,
  ],
  controllers: [BillingController, PublicBillingController, StripeWebhookController],
  exports: [CreditGateService, CreditLedgerService, CreditAccountQueryService],
})
export class CreditModule {}
