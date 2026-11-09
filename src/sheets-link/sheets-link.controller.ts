import { Controller, Delete, Get, UseGuards } from '@nestjs/common';
import { AuthGuard, Session, type AuthUserSession } from '../auth/auth.guard';
import { SheetsLinkService } from './sheets-link.service';

/**
 * Google Sheets add-in linking. Open GET /sheets-link/new-key in a browser where you are
 * signed in to CELLIX, copy the key, paste it into the Sheets sidebar (CELLIX → Settings).
 */
@Controller('sheets-link')
@UseGuards(AuthGuard)
export class SheetsLinkController {
  constructor(private readonly service: SheetsLinkService) {}

  @Get('new-key')
  async newKey(@Session() session: AuthUserSession) {
    const { key, expiresAt } = await this.service.create({
      userId: session.user.id,
      email: session.user.email ?? null,
      name: (session.user as { name?: string | null }).name ?? null,
    });
    return {
      key,
      expiresAt,
      howTo: 'Paste this key into Google Sheets → Extensions → CELLIX → Open → Settings. It is shown only once; create a new one if you lose it.',
    };
  }

  @Delete('keys')
  async revoke(@Session() session: AuthUserSession) {
    return { revoked: await this.service.revokeAll(session.user.id) };
  }
}
