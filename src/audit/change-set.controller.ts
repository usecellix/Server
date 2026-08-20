import {
  Body,
  Controller,
  Get,
  NotFoundException,
  Param,
  Post,
  UnprocessableEntityException,
} from '@nestjs/common';
import { ChangeSetService } from './change-set.service';
import { RevertVerificationError } from './errors/revert-verification.error';

@Controller('audit')
export class ChangeSetController {
  constructor(private readonly changeSetService: ChangeSetService) {}

  @Post('apply/:changeSetId')
  async apply(
    @Param('changeSetId') changeSetId: string,
    // TASKS.md #40/#15 — optional, since most apply calls carry no CONDITIONAL_FORMAT
    // or CREATE_CHART creates at all. `{}` (the frontend's existing empty-body
    // convention) has neither key, so both stay undefined for those calls.
    @Body()
    body?: {
      createdConditionalFormatIds?: { sheetName: string; range: string; ruleId: string }[];
      createdChartIds?: { sheetName: string; sourceRange: string; chartId: string }[];
    },
  ) {
    const changeSet = await this.changeSetService.markApplied(
      changeSetId,
      body?.createdConditionalFormatIds,
      body?.createdChartIds,
    );
    return { changeSet };
  }

  @Post('revert/:changeSetId')
  async revert(@Param('changeSetId') changeSetId: string) {
    try {
      const result = await this.changeSetService.revert(changeSetId);
      return {
        changeSet: result.changeSet,
        inverseActions: result.inverseActions,
      };
    } catch (error) {
      // TASKS.md #19 — fail-closed revert self-verification. 422: the request is
      // well-formed, but the revert cannot be completed without leaving the workbook
      // in a state that doesn't match beforeState.
      if (error instanceof RevertVerificationError) {
        throw new UnprocessableEntityException({
          message: error.message,
          code: error.code,
          changeSetId: error.changeSetId,
          blockingChanges: error.blockingChanges,
        });
      }
      throw error;
    }
  }

  @Get('history/:conversationId')
  async history(@Param('conversationId') conversationId: string) {
    const changeSets = await this.changeSetService.getHistory(conversationId);
    return { changeSets };
  }

  @Get('change-set/:changeSetId')
  async getOne(@Param('changeSetId') changeSetId: string) {
    const changeSet = await this.changeSetService.getById(changeSetId);
    if (!changeSet) {
      throw new NotFoundException(`Change set ${changeSetId} not found`);
    }
    return { changeSet };
  }
}
