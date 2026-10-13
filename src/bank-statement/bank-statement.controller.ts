import { Body, Controller, HttpCode, Post, UseGuards } from '@nestjs/common';
import { AuthGuard } from '../auth/auth.guard';
import { SkipLogCapture } from '../common/decorators/skip-log-capture.decorator';
import { ImportBankStatementDto } from './bank-statement.dto';
import { BankStatementImportResponse, BankStatementService } from './bank-statement.service';

/**
 * Attachment ingestion (Root/ATTACHMENT_EXTRACTION_PLAN.md). The task pane
 * decodes the file and posts the result here, so this is an ordinary JSON
 * route: there is no file upload, and the raw file never reaches the server.
 */
@UseGuards(AuthGuard)
@Controller('ingest')
export class BankStatementController {
  constructor(private readonly bankStatementService: BankStatementService) {}

  /**
   * The request and response are a person's bank transactions, so neither is
   * copied into the request log (`@SkipLogCapture`, and `rawTable` is
   * summarised by `sanitizeLogBody`). The request can also be far larger than
   * Fastify's 1 MiB default; see `ROUTE_BODY_LIMITS`.
   */
  @Post('bank-statement')
  @HttpCode(200)
  @SkipLogCapture()
  importBankStatement(@Body() body: ImportBankStatementDto): BankStatementImportResponse {
    return this.bankStatementService.importStatement(body);
  }
}
