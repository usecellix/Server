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
import {
  WorkflowTrace,
  WorkflowTraceSchema,
} from './schemas/workflow-trace.schema';
import { WorkflowTraceService } from './workflow-trace.service';

@Module({
  imports: [
    LoggerModule.forRootAsync({
      imports: [AppConfigModule],
      inject: [AppConfigService],
      useFactory: (config: AppConfigService) => buildPinoParams(config.nodeEnv),
    }),
    MongooseModule.forFeature([{ name: WorkflowTrace.name, schema: WorkflowTraceSchema }]),
  ],
  controllers: [FrontendLogController],
  providers: [
    RequestFileLoggerService,
    PlannerFileLoggerService,
    FrontendFileLoggerService,
    WorkflowTraceService,
    LogTtlIndexService,
  ],
  exports: [
    LoggerModule,
    RequestFileLoggerService,
    PlannerFileLoggerService,
    FrontendFileLoggerService,
    WorkflowTraceService,
    MongooseModule,
  ],
})
export class LoggingModule {}
