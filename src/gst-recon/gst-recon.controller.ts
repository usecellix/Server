import { Body, Controller, Post } from '@nestjs/common';
import { GstReconcileRequestDto } from './gst-recon.dto';
import { GstReconService } from './gst-recon.service';

@Controller('gst')
export class GstReconController {
  constructor(private readonly gstReconService: GstReconService) {}

  /**
   * Run GST reconciliation in-memory and return summary + Excel actions.
   * POST /gst/reconcile
   */
  @Post('reconcile')
  reconcile(@Body() body: GstReconcileRequestDto) {
    return this.gstReconService.reconcile(body);
  }
}
