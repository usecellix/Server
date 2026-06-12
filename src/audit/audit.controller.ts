import {
  Controller,
  DefaultValuePipe,
  Get,
  ParseIntPipe,
  Query,
} from '@nestjs/common';
import { AuditService } from './audit.service';
import { LLMTier } from '../types/cellix.types';

@Controller('audit')
export class AuditController {
  constructor(private readonly auditService: AuditService) {}

  @Get('logs')
  async getLogs(
    @Query('limit', new DefaultValuePipe(50), ParseIntPipe) limit: number,
    @Query('offset', new DefaultValuePipe(0), ParseIntPipe) offset: number,
    @Query('tier') tier?: string,
    @Query('from') from?: string,
    @Query('to') to?: string,
  ) {
    return this.auditService.getLogs({
      limit,
      offset,
      tier: tier as LLMTier | undefined,
      fromDate: from ? new Date(from) : undefined,
      toDate: to ? new Date(to) : undefined,
    });
  }

  @Get('stats')
  async getStats(
    @Query('from') from?: string,
    @Query('to') to?: string,
  ) {
    const fromDate = from
      ? new Date(from)
      : new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    const toDate = to ? new Date(to) : new Date();
    return this.auditService.getStats(fromDate, toDate);
  }
}
