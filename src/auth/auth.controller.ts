import { Controller, Get, UseGuards } from '@nestjs/common';
import { AuthGuard, Session, type AuthUserSession } from './auth.guard';

@Controller('auth')
export class AuthController {
  @Get('me')
  @UseGuards(AuthGuard)
  getProfile(@Session() session: AuthUserSession) {
    return {
      user: session.user,
      session: {
        id: session.session.id,
        expiresAt: session.session.expiresAt,
      },
    };
  }

  @Get('health')
  authHealth() {
    return { status: 'ok', providers: ['google', 'microsoft'] };
  }
}
