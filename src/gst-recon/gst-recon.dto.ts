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
  Matches,
  MaxLength,
  ValidateNested,
} from 'class-validator';

export const GST_RECON_TYPES = [
  'PR_VS_GSTR2B',
  'PR_VS_GSTR2A',
  'IMS_VS_PR',
  'GSTR3B_VS_GSTR2B',
  'SALES_VS_GSTR1',
] as const;

export type GstReconTypeDto = (typeof GST_RECON_TYPES)[number];

/**
 * Output layout for the casual missed-rows sheet ("Missed vs GSTR-2B" etc.) — never
 * meaningful outside `missed_books_only`, which drives everything else about that sheet.
 * - `categorized` (default): current multi-section layout, grouped by mismatch reason.
 * - `books_flat`: single flat table — every books row NOT found on the portal (any reason).
 * - `portal_flat`: the reverse direction — every portal row NOT found in the books.
 */
export const GST_RECON_LAYOUTS = ['categorized', 'books_flat', 'portal_flat'] as const;
export type GstReconLayoutDto = (typeof GST_RECON_LAYOUTS)[number];

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

  @IsOptional()
  @Allow()
  clientGstin?: string | number;

  @IsOptional()
  @Allow()
  supplyCategory?: string | number;

  @IsOptional()
  @Allow()
  placeOfSupply?: string | number;
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

  /**
   * Books-side register (Purchase Register or Sales Register).
   * For SALES_VS_GSTR1 this holds the sales register grid.
   */
  @IsOptional()
  @ValidateNested()
  @Type(() => SheetPayloadDto)
  purchase_register?: SheetPayloadDto;

  /** Alias for purchase_register — preferred name for sales runs. */
  @IsOptional()
  @ValidateNested()
  @Type(() => SheetPayloadDto)
  books_register?: SheetPayloadDto;

  @IsOptional()
  @ValidateNested()
  @Type(() => SheetPayloadDto)
  portal_file?: SheetPayloadDto;

  /** Extra GSTR-2A grid for purchase recon against 2B and 2A together. */
  @IsOptional()
  @ValidateNested()
  @Type(() => SheetPayloadDto)
  portal_file_2a?: SheetPayloadDto;

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
  financial_year?: string;

  @IsOptional()
  @IsString()
  @MaxLength(200)
  client_name?: string;

  /**
   * Required client taxpayer GSTIN (single run-level key).
   * Deprecated alias: `gstin` — prefer `client_gstin`.
   */
  @IsOptional()
  @IsString()
  @MaxLength(20)
  client_gstin?: string;

  /** @deprecated Use client_gstin */
  @IsOptional()
  @IsString()
  @MaxLength(20)
  gstin?: string;

  @IsOptional()
  @IsString()
  @MaxLength(120)
  operator_name?: string;

  @IsOptional()
  @IsString()
  @MaxLength(120)
  firm_name?: string;

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

  /**
   * Casual chat mode: GSTIN/period optional; sheet write is books rows
   * missing from the portal file(s) only.
   */
  @IsOptional()
  @IsBoolean()
  missed_books_only?: boolean;

  /**
   * Which layout to write the missed_books_only sheet in — see GstReconLayoutDto.
   * Ignored (categorized always applies) for the full, non-casual reconciliation.
   */
  @IsOptional()
  @IsIn([...GST_RECON_LAYOUTS])
  layout?: GstReconLayoutDto;

  /**
   * Inclusive period date range (yyyy-mm-dd) extracted from the chat prompt — when present,
   * books and portal rows outside this range are filtered out before matching runs.
   */
  @IsOptional()
  @IsString()
  @Matches(/^\d{4}-\d{2}-\d{2}$/)
  period_start?: string;

  @IsOptional()
  @IsString()
  @Matches(/^\d{4}-\d{2}-\d{2}$/)
  period_end?: string;

  /** Human-readable period label for chat/report text, e.g. "April 2024". */
  @IsOptional()
  @IsString()
  @MaxLength(60)
  period_label?: string;
}
