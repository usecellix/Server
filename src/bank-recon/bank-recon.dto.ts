import { Type } from 'class-transformer';
import {
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

/**
 * Output layout (each writes to its own distinctly named sheet):
 * - `categorized` (default): BRS on top, then one block per reason.
 * - `books_flat`: ledger entries not found in the bank, one flat table.
 * - `bank_flat`: statement rows not found in the books — the "post these entries" list.
 * - `matched`: full matched list with pass + confidence, for audit trail.
 */
export const BANK_RECON_LAYOUTS = ['categorized', 'books_flat', 'bank_flat', 'matched'] as const;
export type BankReconLayoutDto = (typeof BANK_RECON_LAYOUTS)[number];

export class BankSheetPayloadDto {
  @IsString()
  @MaxLength(200)
  sheet_name!: string;

  @IsOptional()
  @IsNumber()
  headers_row?: number;

  @IsArray()
  data!: unknown[][];

  @IsOptional()
  @IsIn(['XLSX', 'CSV'])
  file_type?: 'XLSX' | 'CSV';
}

export class BankReconSettingsDto {
  @IsOptional()
  @IsNumber()
  clearing_window_days?: number;

  @IsOptional()
  @IsNumber()
  tight_window_days?: number;

  @IsOptional()
  @IsNumber()
  amount_tolerance?: number;

  @IsOptional()
  @IsNumber()
  narration_threshold?: number;

  @IsOptional()
  @IsNumber()
  max_group_size?: number;

  @IsOptional()
  @IsBoolean()
  enable_group_match?: boolean;
}

export class BankReconcileRequestDto {
  @ValidateNested()
  @Type(() => BankSheetPayloadDto)
  books_ledger!: BankSheetPayloadDto;

  @ValidateNested()
  @Type(() => BankSheetPayloadDto)
  bank_statement!: BankSheetPayloadDto;

  /** Previous period's BRS sheet — uncleared items are carried forward into this run. */
  @IsOptional()
  @ValidateNested()
  @Type(() => BankSheetPayloadDto)
  prior_brs?: BankSheetPayloadDto;

  /** ISO dates. When omitted the period is taken from the books ledger's own date range. */
  @IsOptional()
  @Matches(/^\d{4}-\d{2}-\d{2}$/)
  period_start?: string;

  @IsOptional()
  @Matches(/^\d{4}-\d{2}-\d{2}$/)
  period_end?: string;

  @IsOptional()
  @IsString()
  period_label?: string;

  /** Casual runs: no account details needed and validation problems are warnings, not blockers. */
  @IsOptional()
  @IsBoolean()
  casual?: boolean;

  @IsOptional()
  @IsIn(BANK_RECON_LAYOUTS)
  layout?: BankReconLayoutDto;

  /** With `books_flat`: only timing items (uncleared cheques, deposits in transit, stale cheques). */
  @IsOptional()
  @IsIn(['timing'])
  books_filter?: 'timing';

  /** Balances to use when the sheets don't carry them. */
  @IsOptional()
  @IsNumber()
  books_opening_balance?: number;

  @IsOptional()
  @IsNumber()
  bank_opening_balance?: number;

  @IsOptional()
  @IsString()
  account_number?: string;

  @IsOptional()
  @IsString()
  client_name?: string;

  @IsOptional()
  @IsString()
  output_sheet_name?: string;

  @IsOptional()
  @IsObject()
  @ValidateNested()
  @Type(() => BankReconSettingsDto)
  settings?: BankReconSettingsDto;

  @IsOptional()
  @IsString()
  conversation_id?: string;
}

export class BankReconRelayoutRequestDto {
  @IsString()
  layout_cache_id!: string;

  @IsIn(BANK_RECON_LAYOUTS)
  layout!: BankReconLayoutDto;

  @IsOptional()
  @IsString()
  output_sheet_name?: string;
}
