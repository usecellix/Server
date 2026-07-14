import { Module } from '@nestjs/common';
import { AuthController } from './auth.controller';
import { AuthGuard } from './auth.guard';
import { AuthRouteRegistrar } from './auth-route.registrar';

@Module({
  controllers: [AuthController],
  providers: [AuthRouteRegistrar, AuthGuard],
  exports: [AuthGuard],
})
export class AuthModule {}
