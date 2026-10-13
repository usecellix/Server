import { IsArray, IsObject, IsOptional, IsString, MaxLength } from 'class-validator';
import { RawTable } from '../domain-tools/ingestion/raw-table.types';

export class ImportBankStatementDto {
  /**
   * The decoded file. Checked by `assertRawTable` rather than nested
   * class-validator rules: a year of statements is tens of thousands of cells,
   * and instantiating a validated object per cell is slow for no extra safety.
   */
  @IsObject()
  rawTable!: RawTable;

  /** Sheet names already in the workbook, so the new sheet gets a free name. */
  @IsOptional()
  @IsArray()
  @IsString({ each: true })
  existingSheetNames?: string[];

  /** Place the new sheet after this one. */
  @IsOptional()
  @IsString()
  @MaxLength(64)
  activeSheetName?: string;
}
