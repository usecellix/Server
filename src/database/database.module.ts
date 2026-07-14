import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { AppConfigService } from '../config/app-config.service';

@Module({
  imports: [
    MongooseModule.forRootAsync({
      imports: [AppConfigModule],
      inject: [AppConfigService],
      useFactory: (config: AppConfigService) => {
        const uri = config.mongoUrl.includes('retryWrites=')
          ? config.mongoUrl
          : config.mongoUrl.includes('?')
            ? `${config.mongoUrl}&retryWrites=false`
            : `${config.mongoUrl}?retryWrites=false`;
        return {
          uri,
          dbName: config.mongoDbName,
          retryWrites: false,
        };
      },
    }),
  ],
})
export class DatabaseModule {}
