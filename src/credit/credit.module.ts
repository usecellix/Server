import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { CreditAccount, CreditAccountSchema } from './schemas/credit-account.schema';
import { CreditLedgerEntry, CreditLedgerEntrySchema } from './schemas/credit-ledger.schema';
import {
  ProcessedRazorpayEvent,
  ProcessedRazorpayEventSchema,
  Subscription,
  SubscriptionSchema,
} from './schemas/subscription.schema';
import { CreditGateService } from './credit-gate.service';
import { CreditLedgerService } from './credit-ledger.service';
import { CreditAccountQueryService } from './credit-account-query.service';
import { RazorpayCheckoutService } from './razorpay-checkout.service';
import { RazorpayWebhookService } from './razorpay-webhook.service';
import { BillingController, PublicBillingController, RazorpayWebhookController } from './billing.controller';

@Module({
  imports: [
    AppConfigModule,
    MongooseModule.forFeature([
      { name: CreditAccount.name, schema: CreditAccountSchema },
      { name: CreditLedgerEntry.name, schema: CreditLedgerEntrySchema },
      { name: Subscription.name, schema: SubscriptionSchema },
      { name: ProcessedRazorpayEvent.name, schema: ProcessedRazorpayEventSchema },
    ]),
  ],
  providers: [
    CreditGateService,
    CreditLedgerService,
    CreditAccountQueryService,
    RazorpayCheckoutService,
    RazorpayWebhookService,
  ],
  controllers: [BillingController, PublicBillingController, RazorpayWebhookController],
  exports: [CreditGateService, CreditLedgerService, CreditAccountQueryService],
})
export class CreditModule {}
