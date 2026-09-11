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
        // Do NOT force retryWrites=false here. That was correct only for a local
        // standalone mongod (which rejects retryable writes without a replica
        // set) and was previously hardcoded onto every connection string,
        // silently disabling retry-on-transient-network-blip even against a
        // replica set / Atlas, where MongoDB's own default is retryWrites=true
        // and disabling it is actively worse. Let the driver's default (true)
        // apply; a caller who genuinely needs it off can still say so in
        // MONGODB_URL (e.g. a standalone dev mongod without a replica set).
        return {
          uri: config.mongoUrl,
          dbName: config.mongoDbName,
        };
      },
    }),
  ],
})
export class DatabaseModule {}
