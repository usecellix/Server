import { Body, Controller, Param, Post } from '@nestjs/common';
import {
  IsArray,
  IsIn,
  IsObject,
  IsOptional,
  IsString,
  ValidateNested,
} from 'class-validator';
import { Type } from 'class-transformer';
import { GstReconcileRequestDto } from './gst-recon.dto';
import { GstReconService } from './gst-recon.service';
import { GstReconOrchestrator } from './orchestrator';
import { GstReconIntent, GstReconIntentPayload } from './orchestrator.types';

class SheetHeaderDto {
  @IsString()
  sheetName!: string;

  @IsArray()
  headers!: string[];
}

class ConversationalReconDto {
  @IsIn(['GST_RECON_PURCHASE', 'GST_RECON_SALES'])
  intent!: GstReconIntent;

  @IsOptional()
  @IsString()
  extractedClientName?: string;

  @IsOptional()
  @IsString()
  extractedGstin?: string;

  @IsOptional()
  @IsString()
  extractedPeriod?: string;

  @IsOptional()
  @IsString()
  extractedFinancialYear?: string;

  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => SheetHeaderDto)
  sheets!: SheetHeaderDto[];

  @IsOptional()
  @IsObject()
  sheetData?: Record<string, unknown[][]>;
}

class AuditOutcomeDto {
  @IsIn(['applied', 'rejected'])
  outcome!: 'applied' | 'rejected';
}

@Controller('gst')
export class GstReconController {
  constructor(
    private readonly gstReconService: GstReconService,
    private readonly orchestrator: GstReconOrchestrator,
  ) {}

  @Post('reconcile')
  reconcile(@Body() body: GstReconcileRequestDto) {
    return this.gstReconService.reconcile(body);
  }

  @Post('reconcile/conversational')
  conversational(@Body() body: ConversationalReconDto) {
    const payload: GstReconIntentPayload = {
      intent: body.intent,
      extractedClientName: body.extractedClientName,
      extractedGstin: body.extractedGstin,
      extractedPeriod: body.extractedPeriod,
      extractedFinancialYear: body.extractedFinancialYear,
    };
    return this.orchestrator.runConversationalRecon(
      body.intent,
      payload,
      body.sheets,
      body.sheetData,
    );
  }

  @Post('reconcile/audit/:id/outcome')
  auditOutcome(@Param('id') id: string, @Body() body: AuditOutcomeDto) {
    return this.gstReconService.updateAuditOutcome(id, body.outcome);
  }
}
