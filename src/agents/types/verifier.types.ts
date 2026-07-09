import { SubTask } from './agent.types';

export interface VerificationResult {
  passed: boolean;
  failedStepIndex?: number;
  failedStepId?: string;
  issue?: string;
  correction?: string;
  rawCellValue?: unknown;
  expectedCellValue?: unknown;
}

export interface StepRetryContext {
  originalStep: SubTask;
  attempt: number;
  maxAttempts: number;
  verifierFeedback: string;
}
