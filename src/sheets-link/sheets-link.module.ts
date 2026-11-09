import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { SheetsLinkController } from './sheets-link.controller';
import { SheetsLinkKey, SheetsLinkKeySchema } from './sheets-link.schema';
import { SheetsLinkService } from './sheets-link.service';

@Module({
  imports: [AppConfigModule, MongooseModule.forFeature([{ name: SheetsLinkKey.name, schema: SheetsLinkKeySchema }])],
  providers: [SheetsLinkService],
  controllers: [SheetsLinkController],
})
export class SheetsLinkModule {}
