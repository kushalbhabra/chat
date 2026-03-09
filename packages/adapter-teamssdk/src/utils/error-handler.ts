/**
 * Error codes used by the Teams adapter.
 */
export enum TeamsAdapterErrorCode {
  UNAUTHORIZED = 'UNAUTHORIZED',
  FORBIDDEN = 'FORBIDDEN',
  NOT_FOUND = 'NOT_FOUND',
  RATE_LIMITED = 'RATE_LIMITED',
  VALIDATION_ERROR = 'VALIDATION_ERROR',
  API_ERROR = 'API_ERROR',
  UNKNOWN = 'UNKNOWN',
}

/**
 * Typed error thrown by the Teams adapter.
 */
export class TeamsAdapterError extends Error {
  readonly code: TeamsAdapterErrorCode;
  readonly statusCode: number;
  readonly originalError?: unknown;

  constructor(
    message: string,
    code: TeamsAdapterErrorCode = TeamsAdapterErrorCode.UNKNOWN,
    statusCode = 500,
    originalError?: unknown
  ) {
    super(message);
    this.name = 'TeamsAdapterError';
    this.code = code;
    this.statusCode = statusCode;
    this.originalError = originalError;
  }
}

interface HttpErrorBody {
  error?: { message?: string; code?: string };
  message?: string;
}

/**
 * Maps an HTTP fetch error (or any unknown thrown value) to a TeamsAdapterError.
 */
export function handleApiError(error: unknown): TeamsAdapterError {
  if (error instanceof TeamsAdapterError) {
    return error;
  }

  if (error instanceof Response) {
    return mapHttpResponse(error, undefined);
  }

  if (
    typeof error === 'object' &&
    error !== null &&
    'status' in error &&
    typeof (error as Record<string, unknown>)['status'] === 'number'
  ) {
    const httpErr = error as { status: number; statusText?: string; body?: HttpErrorBody };
    const message =
      httpErr.body?.error?.message ??
      httpErr.body?.message ??
      httpErr.statusText ??
      'Unknown HTTP error';
    return mapHttpResponse({ status: httpErr.status } as Response, message);
  }

  if (error instanceof Error) {
    return new TeamsAdapterError(error.message, TeamsAdapterErrorCode.API_ERROR, 500, error);
  }

  return new TeamsAdapterError(
    typeof error === 'string' ? error : 'An unknown error occurred',
    TeamsAdapterErrorCode.UNKNOWN,
    500,
    error
  );
}

/**
 * Creates a TeamsAdapterError from a failed HTTP Response object.
 */
export async function handleHttpError(response: Response): Promise<TeamsAdapterError> {
  let body: HttpErrorBody | undefined;
  try {
    body = (await response.json()) as HttpErrorBody;
  } catch {
    // Ignore body parse failure
  }
  const message =
    body?.error?.message ?? body?.message ?? response.statusText ?? 'HTTP error';
  return mapHttpResponse(response, message);
}

function mapHttpResponse(response: Pick<Response, 'status'>, message: string | undefined): TeamsAdapterError {
  const msg = message ?? 'HTTP error';
  switch (response.status) {
    case 400:
      return new TeamsAdapterError(msg, TeamsAdapterErrorCode.VALIDATION_ERROR, 400);
    case 401:
      return new TeamsAdapterError(msg, TeamsAdapterErrorCode.UNAUTHORIZED, 401);
    case 403:
      return new TeamsAdapterError(msg, TeamsAdapterErrorCode.FORBIDDEN, 403);
    case 404:
      return new TeamsAdapterError(msg, TeamsAdapterErrorCode.NOT_FOUND, 404);
    case 429:
      return new TeamsAdapterError(msg, TeamsAdapterErrorCode.RATE_LIMITED, 429);
    default:
      return new TeamsAdapterError(msg, TeamsAdapterErrorCode.API_ERROR, response.status);
  }
}

/**
 * Returns true if the error is a rate-limit error (HTTP 429).
 */
export function isRateLimitError(error: unknown): error is TeamsAdapterError {
  return (
    error instanceof TeamsAdapterError &&
    error.code === TeamsAdapterErrorCode.RATE_LIMITED
  );
}

/**
 * Extracts the Retry-After delay in milliseconds from a rate-limit error or
 * a Response headers object. Falls back to the provided default.
 */
export function getRateLimitDelay(
  headersOrError: unknown,
  defaultDelayMs = 1000
): number {
  if (headersOrError instanceof Headers) {
    const retryAfter = headersOrError.get('retry-after');
    if (retryAfter) {
      const seconds = Number(retryAfter);
      if (!isNaN(seconds)) return seconds * 1000;
    }
  }
  return defaultDelayMs;
}
