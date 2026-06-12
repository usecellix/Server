import { Allow, IsString } from 'class-validator';
import { WorkbookContext } from '../../types/cellix.types';

export class CompareSheetsRequestDto {
  @IsString()
  sheetA!: string;

  @IsString()
  sheetB!: string;

  @Allow()
  context!: WorkbookContext;
}
