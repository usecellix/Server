import { Body, Controller, Get, Param, Post, UnprocessableEntityException } from '@nestjs/common';
import { CheckpointService } from './checkpoint.service';
import { RestoreVerificationError } from './errors/restore-verification.error';

interface CreateCheckpointBody {
  workbookId: string;
  conversationId: string;
  label?: string;
}

@Controller('audit')
export class CheckpointController {
  constructor(private readonly checkpointService: CheckpointService) {}

  @Post('checkpoint')
  async create(@Body() body: CreateCheckpointBody) {
    const checkpoint = await this.checkpointService.createManual({
      workbookId: body.workbookId,
      conversationId: body.conversationId,
      label: body.label,
    });
    return { checkpoint };
  }

  @Get('checkpoint/:workbookId')
  async list(@Param('workbookId') workbookId: string) {
    const checkpoints = await this.checkpointService.listByWorkbook(workbookId);
    return { checkpoints };
  }

  @Post('checkpoint/restore/:checkpointId')
  async restore(@Param('checkpointId') checkpointId: string) {
    try {
      const result = await this.checkpointService.restore(checkpointId);
      return result;
    } catch (error) {
      // TASKS.md #29 — fail-closed restore. 422: well-formed request, but the
      // restore cannot be completed without leaving the workbook in a state
      // that doesn't match the checkpoint's anchor.
      if (error instanceof RestoreVerificationError) {
        throw new UnprocessableEntityException({
          message: error.message,
          code: error.code,
          checkpointId: error.checkpointId,
          blockingChangeSetId: error.blockingChangeSetId,
        });
      }
      throw error;
    }
  }
}
