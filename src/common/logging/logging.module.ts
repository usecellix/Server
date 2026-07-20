import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { LoggerModule } from 'nestjs-pino';
import { AppConfigModule } from '../../config/app-config.module';
import { AppConfigService } from '../../config/app-config.service';
import { buildPinoParams } from './pino-config';
import { FrontendFileLoggerService } from './frontend-file-logger.service';
import { FrontendLogController } from './frontend-log.controller';
import { LogTtlIndexService } from './log-ttl-index.service';
import { PlannerFileLoggerService } from './planner-file-logger.service';
import { RequestFileLoggerService } from './request-file-logger.service';
import { FrontendLog, FrontendLogSchema } from './schemas/frontend-log.schema';
import { PlannerLog, PlannerLogSchema } from './schemas/planner-log.schema';
import { RequestLog, RequestLogSchema } from './schemas/request-log.schema';

@Module({
  imports: [
    LoggerModule.forRootAsync({
      imports: [AppConfigModule],
      inject: [AppConfigService],
      useFactory: (config: AppConfigService) => buildPinoParams(config.nodeEnv),
    }),
    MongooseModule.forFeature([
      { name: RequestLog.name, schema: RequestLogSchema },
      { name: PlannerLog.name, schema: PlannerLogSchema },
      { name: FrontendLog.name, schema: FrontendLogSchema },
    ]),
  ],
  controllers: [FrontendLogController],
  providers: [
    RequestFileLoggerService,
    PlannerFileLoggerService,
    FrontendFileLoggerService,
    LogTtlIndexService,
  ],
  exports: [
    LoggerModule,
    RequestFileLoggerService,
    PlannerFileLoggerService,
    FrontendFileLoggerService,
    MongooseModule,
  ],
})
export class LoggingModule {}
