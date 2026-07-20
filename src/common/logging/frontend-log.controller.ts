import { Body, Controller, Headers, Post } from '@nestjs/common';
import { FrontendLogBatchDto } from './dto/frontend-log-batch.dto';
import { FrontendFileLoggerService } from './frontend-file-logger.service';

@Controller('telemetry')
export class FrontendLogController {
  constructor(private readonly frontendFileLogger: FrontendFileLoggerService) {}

  /**
   * Ingest Excel add-in client events (console errors, Accept/Reject, preview apply).
   * Fire-and-forget from the client — always returns quickly.
   */
  @Post('frontend')
  ingest(
    @Body() body: FrontendLogBatchDto,
    @Headers('user-agent') userAgent?: string,
  ): { accepted: number } {
    const ua = typeof userAgent === 'string' ? userAgent.slice(0, 500) : undefined;
    this.frontendFileLogger.logEvents(
      body.events.map((event) => ({
        ...event,
        userAgent: event.userAgent ?? ua,
      })),
    );
    return { accepted: body.events.length };
  }
}
