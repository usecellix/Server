import { CallHandler, ExecutionContext, Injectable, NestInterceptor } from '@nestjs/common';
import { Observable, tap } from 'rxjs';
import {
  captureJsonResponse,
  captureTextResponse,
  type ReplyWithRaw,
} from '../logging/request-response-capture.util';

/**
 * Stores JSON/text handler results on the reply so the request file logger
 * can include them. SSE bodies are captured separately in sse.util.
 */
@Injectable()
export class RequestResponseCaptureInterceptor implements NestInterceptor {
  intercept(context: ExecutionContext, next: CallHandler): Observable<unknown> {
    if (context.getType() !== 'http') {
      return next.handle();
    }

    const reply = context.switchToHttp().getResponse<ReplyWithRaw>();

    return next.handle().pipe(
      tap({
        next: (data) => {
          if (typeof data === 'string') {
            captureTextResponse(reply, data);
          } else if (data !== undefined) {
            captureJsonResponse(reply, data);
          }
        },
      }),
    );
  }
}
