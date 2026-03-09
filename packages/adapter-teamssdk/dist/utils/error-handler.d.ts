/**
 * Error codes used by the Teams adapter.
 */
export declare enum TeamsAdapterErrorCode {
    UNAUTHORIZED = "UNAUTHORIZED",
    FORBIDDEN = "FORBIDDEN",
    NOT_FOUND = "NOT_FOUND",
    RATE_LIMITED = "RATE_LIMITED",
    VALIDATION_ERROR = "VALIDATION_ERROR",
    API_ERROR = "API_ERROR",
    UNKNOWN = "UNKNOWN"
}
/**
 * Typed error thrown by the Teams adapter.
 */
export declare class TeamsAdapterError extends Error {
    readonly code: TeamsAdapterErrorCode;
    readonly statusCode: number;
    readonly originalError?: unknown;
    constructor(message: string, code?: TeamsAdapterErrorCode, statusCode?: number, originalError?: unknown);
}
/**
 * Maps an HTTP fetch error (or any unknown thrown value) to a TeamsAdapterError.
 */
export declare function handleApiError(error: unknown): TeamsAdapterError;
/**
 * Creates a TeamsAdapterError from a failed HTTP Response object.
 */
export declare function handleHttpError(response: Response): Promise<TeamsAdapterError>;
/**
 * Returns true if the error is a rate-limit error (HTTP 429).
 */
export declare function isRateLimitError(error: unknown): error is TeamsAdapterError;
/**
 * Extracts the Retry-After delay in milliseconds from a rate-limit error or
 * a Response headers object. Falls back to the provided default.
 */
export declare function getRateLimitDelay(headersOrError: unknown, defaultDelayMs?: number): number;
//# sourceMappingURL=error-handler.d.ts.map