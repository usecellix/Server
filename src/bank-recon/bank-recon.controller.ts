import { Body, Controller, Post } from '@nestjs/common';
import { BankReconRelayoutRequestDto, BankReconcileRequestDto } from './bank-recon.dto';
import { BankReconService } from './bank-recon.service';

@Controller('bank')
export class BankReconController {
  constructor(private readonly bankReconService: BankReconService) {}

  @Post('reconcile')
  reconcile(@Body() body: BankReconcileRequestDto) {
    return this.bankReconService.reconcile(body);
  }

  /** Follow-up to an earlier run — "change the layout" re-maps cached results. */
  @Post('reconcile/relayout')
  relayout(@Body() body: BankReconRelayoutRequestDto) {
    return this.bankReconService.relayout(body);
  }
}
