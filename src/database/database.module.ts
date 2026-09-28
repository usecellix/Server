import { Logger, Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import type { Connection } from 'mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { AppConfigService } from '../config/app-config.service';

const logger = new Logger('DatabaseModule');

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
          // Fires once the underlying connection actually opens — the useFactory
          // return value above only configures the connection attempt, it runs
          // before Mongoose has connected. `connection.on('connected', ...)` also
          // covers a driver-level reconnect after a dropped connection, which a
          // one-shot post-bootstrap log would miss.
          connectionFactory: (connection: Connection) => {
            const logConnected = () =>
              logger.log(`MongoDB connected: db="${connection.db?.databaseName ?? config.mongoDbName}"`);
            // readyState can already be 1 (connected) by the time this factory
            // runs — Mongoose connects as soon as `mongoose.createConnection`
            // is called, which happens before this callback for a fast local
            // connection, so a plain `.on('connected', ...)` can miss the event
            // entirely. Check the current state first, still listen for future
            // (re)connects.
            if (connection.readyState === 1) {
              logConnected();
            } else {
              connection.on('connected', logConnected);
            }
            connection.on('error', (err: Error) => {
              logger.error(`MongoDB connection error: ${err.message}`, err.stack);
            });
            return connection;
          },
        };
      },
    }),
  ],
})
export class DatabaseModule {}
