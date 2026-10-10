import { SetMetadata } from '@nestjs/common';

export const SKIP_LOG_CAPTURE_KEY = 'skipLogCapture';

/**
 * Keeps a route's response body out of `logs/requests.log`. For routes whose
 * response is the user's own financial data rather than app output.
 */
export const SkipLogCapture = () => SetMetadata(SKIP_LOG_CAPTURE_KEY, true);
