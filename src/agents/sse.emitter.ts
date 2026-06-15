import { Action } from './types/agent.types';

export type SseEventPayload =
  | { type: 'THINKING'; message: string }
  | { type: 'CLARIFY'; questions: string[] }
  | { type: 'CHECKPOINT'; step: string }
  | { type: 'ACTION'; action: Action }
  | { type: 'VERIFY_PASS' }
  | { type: 'VERIFY_FAIL'; feedback: string }
  | { type: 'DONE'; summary: string }
  | { type: 'ERROR'; message: string };

export type SseEmitFn = (event: string, data: Record<string, unknown>) => void;

export class SseEmitter {
  constructor(private readonly emit: SseEmitFn) {}

  send(payload: SseEventPayload): void {
    switch (payload.type) {
      case 'THINKING':
        this.emit('thinking', { message: payload.message });
        break;
      case 'CLARIFY':
        this.emit('clarification', {
          question: payload.questions[0] ?? 'Need more information',
          suggestions: payload.questions.slice(1),
          ambiguityScore: 75,
        });
        break;
      case 'CHECKPOINT':
        this.emit('status', { message: payload.step });
        break;
      case 'ACTION':
        this.emit('status', { message: `Prepared ${payload.action.type} action` });
        break;
      case 'VERIFY_PASS':
        this.emit('status', { message: 'Actions verified' });
        break;
      case 'VERIFY_FAIL':
        this.emit('status', { message: `Verification issue: ${payload.feedback}` });
        break;
      case 'DONE':
        this.emit('conversation_end', { summary: payload.summary });
        break;
      case 'ERROR':
        this.emit('error', { message: payload.message });
        break;
      default:
        break;
    }
  }
}
