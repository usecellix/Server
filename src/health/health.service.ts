import { Injectable } from '@nestjs/common';

@Injectable()
export class HealthService {
  getStatus(): { status: string; message: string } {
    return {
      status: 'ok',
      message: 'NestJS server is running',
    };
  }
}
