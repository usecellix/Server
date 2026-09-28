import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { CreditAccount, CreditAccountSchema } from '../credit/schemas/credit-account.schema';
import { CreditLedgerEntry, CreditLedgerEntrySchema } from '../credit/schemas/credit-ledger.schema';
import { Subscription, SubscriptionSchema } from '../credit/schemas/subscription.schema';
import { AdminController } from './admin.controller';
import { AdminGuard } from './admin.guard';
import { AdminUsersService } from './admin-users.service';

@Module({
  imports: [
    AppConfigModule,
    MongooseModule.forFeature([
      { name: CreditAccount.name, schema: CreditAccountSchema },
      { name: CreditLedgerEntry.name, schema: CreditLedgerEntrySchema },
      { name: Subscription.name, schema: SubscriptionSchema },
    ]),
  ],
  controllers: [AdminController],
  providers: [AdminGuard, AdminUsersService],
})
export class AdminModule {}
