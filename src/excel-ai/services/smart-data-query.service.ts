import { Injectable, Logger } from '@nestjs/common';
import { WorkbookContext } from '../../types/cellix.types';
import {
  buildDataQuerySystemPrompt,
  buildDataQueryUserMessage,
} from '../prompts/data-query-system-prompt';
import { sliceRelevantColumns } from '../utils/column-slicer.util';
import { OpenRouterService } from './openrouter.service';

export type SmartDataQueryEmit = (event: string, data: Record<string, unknown>) => void;

@Injectable()
export class SmartDataQueryService {
  private readonly logger = new Logger(SmartDataQueryService.name);

  constructor(private readonly openRouter: OpenRouterService) {}

  /**
   * Answer a read-only data query using column-sliced sheet data and the MEDIUM LLM tier.
   */
  async handleQuery(
    message: string,
    sheetData: unknown[][],
    workbookContext: WorkbookContext | undefined,
    activeSheetName: string | undefined,
    emit: SmartDataQueryEmit,
  ): Promise<string> {
    const sliceResult = sliceRelevantColumns(
      message,
      workbookContext,
      sheetData,
      activeSheetName,
    );

    if (!sliceResult.sheets.length || !sliceResult.sheets[0].rows.length) {
      this.logger.warn('SmartDataQuery: no sheet data available');
      return 'I could not find any sheet data to answer your question. Please make sure a sheet with data is active.';
    }

    const sheet = sliceResult.sheets[0];

    this.logger.log(
      `SmartDataQuery: sheet=${sheet.sheetName} cols=${sheet.headers.join(',')} rows=${sheet.totalRows}`,
    );

    emit('thinking', {
      message: `Reading ${sheet.sheetName} and analyzing ${sheet.headers.join(', ')}`,
    });

    const systemPrompt = buildDataQuerySystemPrompt();
    const userMessage = buildDataQueryUserMessage(message, sheet);

    try {
      const answer = await this.openRouter.complete({
        systemPrompt,
        userMessage,
        tier: 'medium',
        maxTokens: 512,
        responseFormat: 'text',
        reasoningEffort: 'none',
      });
      return answer.trim();
    } catch (error) {
      this.logger.error('SmartDataQuery LLM error', error);
      return 'I was unable to compute the answer from the sheet data. Please try again.';
    }
  }
}
