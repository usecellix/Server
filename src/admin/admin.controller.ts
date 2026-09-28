import { Controller, Get, NotFoundException, Param, Query, UseGuards } from '@nestjs/common';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { AdminGuard } from './admin.guard';
import { AdminUsersService, UserSort } from './admin-users.service';

const VALID_SORTS: UserSort[] = ['recent', 'spend', 'prompts', 'credits', 'creditsUsed', 'seen'];

/**
 * Read surface for the Dashboard admin app — the reference implementation of
 * routing the Dashboard through the backend instead of a direct MongoClient
 * (see AdminGuard's docblock). Only `/admin/users*` is migrated so far;
 * prompts/billing/overview/models stay on the Dashboard's own MongoClient
 * until they're moved the same way.
 */
@UseGuards(AdminGuard)
@Controller('admin')
export class AdminController {
  constructor(private readonly adminUsers: AdminUsersService) {}

  @Get('users')
  @SkipEnvelope()
  async listUsers(
    @Query('q') q?: string,
    @Query('page') page?: string,
    @Query('sort') sort?: string,
    @Query('plan') plan?: string,
    @Query('sinceDays') sinceDays?: string,
  ) {
    const resolvedSort = VALID_SORTS.includes(sort as UserSort) ? (sort as UserSort) : 'recent';
    const days = Number(sinceDays);
    return this.adminUsers.listUsers({
      q,
      plan,
      sort: resolvedSort,
      page: Math.max(1, Number(page) || 1),
      since: Number.isFinite(days) && days > 0 ? new Date(Date.now() - days * 86_400_000) : undefined,
    });
  }

  @Get('users/:id')
  @SkipEnvelope()
  async getUser(@Param('id') id: string) {
    const user = await this.adminUsers.getUser(id);
    if (!user) throw new NotFoundException('USER_NOT_FOUND');
    return user;
  }
}
