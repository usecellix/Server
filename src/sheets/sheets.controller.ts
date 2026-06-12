import { Body, Controller, HttpException, HttpStatus, Post } from '@nestjs/common';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { CompareSheetsRequestDto } from './dto/compare-sheets-request.dto';
import { MultiSheetService } from './multi-sheet.service';

@Controller('sheets')
export class SheetsController {
  constructor(private readonly multiSheetService: MultiSheetService) {}

  @Post('compare')
  @SkipEnvelope()
  async compareSheets(@Body() body: CompareSheetsRequestDto) {
    const { sheetA, sheetB, context } = body;

    const snapshotA = context.sheets.find((sheet) => sheet.sheetName === sheetA);
    const snapshotB = context.sheets.find((sheet) => sheet.sheetName === sheetB);

    if (!snapshotA) {
      throw new HttpException(`Sheet "${sheetA}" not found in context`, HttpStatus.BAD_REQUEST);
    }
    if (!snapshotB) {
      throw new HttpException(`Sheet "${sheetB}" not found in context`, HttpStatus.BAD_REQUEST);
    }

    return this.multiSheetService.compareSheets(snapshotA, snapshotB);
  }
}
