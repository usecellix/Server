import { Body, Controller, Post, UseGuards } from '@nestjs/common';
import { AuthGuard } from '../auth/auth.guard';
import { SkipLogCapture } from '../common/decorators/skip-log-capture.decorator';
import { BankReconRelayoutRequestDto, BankReconcileRequestDto } from './bank-recon.dto';
import { BankReconService } from './bank-recon.service';

/**
 * Both routes carry a person's bank transactions: signed-in users only, and neither the request
 * nor the response is copied into the request log.
 */
@UseGuards(AuthGuard)
@Controller('bank')
export class BankReconController {
  constructor(private readonly bankReconService: BankReconService) {}

  @Post('reconcile')
  @SkipLogCapture()
  reconcile(@Body() body: BankReconcileRequestDto) {
    return this.bankReconService.reconcile(body);
  }

  /** Follow-up to an earlier run — "change the layout" re-maps cached results. */
  @Post('reconcile/relayout')
  @SkipLogCapture()
  relayout(@Body() body: BankReconRelayoutRequestDto) {
    return this.bankReconService.relayout(body);
  }
}
