import {
  Controller,
  Get,
  NotFoundException,
  Param,
  Post,
} from '@nestjs/common';
import { ChangeSetService } from './change-set.service';

@Controller('audit')
export class ChangeSetController {
  constructor(private readonly changeSetService: ChangeSetService) {}

  @Post('apply/:changeSetId')
  async apply(@Param('changeSetId') changeSetId: string) {
    const changeSet = await this.changeSetService.markApplied(changeSetId);
    return { changeSet };
  }

  @Post('revert/:changeSetId')
  async revert(@Param('changeSetId') changeSetId: string) {
    const result = await this.changeSetService.revert(changeSetId);
    return {
      changeSet: result.changeSet,
      inverseActions: result.inverseActions,
    };
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
