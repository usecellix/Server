import { Module } from '@nestjs/common';
import { AppConfigModule } from '../config/app-config.module';
import { AuthController, SignInProvidersController } from './auth.controller';
import { AuthGuard } from './auth.guard';
import { AuthRouteRegistrar } from './auth-route.registrar';
import { ExcelLoginBridgeService } from './excel-login-bridge.service';
import { ExcelLoginController } from './excel-login.controller';

@Module({
  imports: [AppConfigModule],
  controllers: [AuthController, SignInProvidersController, ExcelLoginController],
  providers: [AuthRouteRegistrar, AuthGuard, ExcelLoginBridgeService],
  exports: [AuthGuard],
})
export class AuthModule {}
