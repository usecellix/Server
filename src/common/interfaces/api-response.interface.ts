export interface ApiSuccessResponse<T> {
  success: true;
  traceId: string;
  data: T;
}

export interface ApiErrorResponse {
  success: false;
  traceId: string;
  error: {
    code: string;
    message: string;
    details?: unknown;
  };
}
