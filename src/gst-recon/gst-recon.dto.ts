import { Type } from 'class-transformer';
import {
  Allow,
  IsArray,
  IsBoolean,
  IsIn,
  IsNumber,
  IsObject,
  IsOptional,
  IsString,
  MaxLength,
  ValidateNested,
} from 'class-validator';

export const GST_RECON_TYPES = [
  'PR_VS_GSTR2B',
  'PR_VS_GSTR2A',
  'IMS_VS_PR',
  'GSTR3B_VS_GSTR2B',
] as const;

export type GstReconTypeDto = (typeof GST_RECON_TYPES)[number];

export class ColumnMappingDto {
  @IsOptional()
  @Allow()
  gstin?: string | number;

  @IsOptional()
  @Allow()
  invoiceNo?: string | number;

  @IsOptional()
  @Allow()
  invoice_no?: string | number;

  @IsOptional()
  @Allow()
  invoiceDate?: string | number;

  @IsOptional()
  @Allow()
  taxableAmt?: string | number;

  @IsOptional()
  @Allow()
  taxable_amt?: string | number;

  @IsOptional()
  @Allow()
  taxAmount?: string | number;

  @IsOptional()
  @Allow()
  igst?: string | number;

  @IsOptional()
  @Allow()
  cgst?: string | number;

  @IsOptional()
  @Allow()
  sgst?: string | number;

  @IsOptional()
  @Allow()
  narration?: string | number;

  @IsOptional()
  @Allow()
  irn?: string | number;

  @IsOptional()
  @Allow()
  documentType?: string | number;

  @IsOptional()
  @Allow()
  imsAction?: string | number;
}

export class SheetPayloadDto {
  @IsString()
  @MaxLength(200)
  sheet_name!: string;

  @IsOptional()
  @IsNumber()
  headers_row?: number;

  @IsArray()
  data!: unknown[][];

  @IsOptional()
  @IsObject()
  @ValidateNested()
  @Type(() => ColumnMappingDto)
  column_mapping?: ColumnMappingDto;

  @IsOptional()
  @IsString()
  file_type?: string;
}

export class GstReconSettingsDto {
  @IsOptional()
  @IsNumber()
  amount_tolerance_abs?: number;

  @IsOptional()
  @IsNumber()
  amount_tolerance_pct?: number;

  @IsOptional()
  @IsNumber()
  invoice_fuzzy_threshold?: number;

  @IsOptional()
  @IsNumber()
  date_tolerance_days?: number;

  @IsOptional()
  @IsBoolean()
  detect_rcm?: boolean;

  @IsOptional()
  @IsBoolean()
  use_ims_data?: boolean;
}

export class Gstr3bComponentDto {
  @IsNumber()
  igst!: number;

  @IsNumber()
  cgst!: number;

  @IsNumber()
  sgst!: number;
}

export class GstReconcileRequestDto {
  @IsIn([...GST_RECON_TYPES])
  reconciliation_type!: GstReconTypeDto;

  @IsOptional()
  @ValidateNested()
  @Type(() => SheetPayloadDto)
  purchase_register?: SheetPayloadDto;

  @IsOptional()
  @ValidateNested()
  @Type(() => SheetPayloadDto)
  portal_file?: SheetPayloadDto;

  @IsOptional()
  @ValidateNested()
  @Type(() => SheetPayloadDto)
  ims_data?: SheetPayloadDto;

  @IsOptional()
  @ValidateNested()
  @Type(() => GstReconSettingsDto)
  settings?: GstReconSettingsDto;

  @IsOptional()
  @IsString()
  @MaxLength(20)
  period?: string;

  @IsOptional()
  @IsString()
  @MaxLength(20)
  gstin?: string;

  @IsOptional()
  @IsString()
  session_id?: string;

  @IsOptional()
  @IsString()
  conversation_id?: string;

  /** For GSTR3B_VS_GSTR2B summary recon */
  @IsOptional()
  @ValidateNested()
  @Type(() => Gstr3bComponentDto)
  gstr3b_itc?: Gstr3bComponentDto;

  @IsOptional()
  @ValidateNested()
  @Type(() => Gstr3bComponentDto)
  gstr2b_itc?: Gstr3bComponentDto;

  @IsOptional()
  @IsString()
  @MaxLength(120)
  output_sheet_name?: string;
}
