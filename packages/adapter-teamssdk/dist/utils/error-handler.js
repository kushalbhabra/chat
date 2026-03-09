/**
 * Error codes used by the Teams adapter.
 */
export var TeamsAdapterErrorCode;
(function (TeamsAdapterErrorCode) {
    TeamsAdapterErrorCode["UNAUTHORIZED"] = "UNAUTHORIZED";
    TeamsAdapterErrorCode["FORBIDDEN"] = "FORBIDDEN";
    TeamsAdapterErrorCode["NOT_FOUND"] = "NOT_FOUND";
    TeamsAdapterErrorCode["RATE_LIMITED"] = "RATE_LIMITED";
    TeamsAdapterErrorCode["VALIDATION_ERROR"] = "VALIDATION_ERROR";
    TeamsAdapterErrorCode["API_ERROR"] = "API_ERROR";
    TeamsAdapterErrorCode["UNKNOWN"] = "UNKNOWN";
})(TeamsAdapterErrorCode || (TeamsAdapterErrorCode = {}));
/**
 * Typed error thrown by the Teams adapter.
 */
export class TeamsAdapterError extends Error {
    code;
    statusCode;
    originalError;
    constructor(message, code = TeamsAdapterErrorCode.UNKNOWN, statusCode = 500, originalError) {
        super(message);
        this.name = 'TeamsAdapterError';
        this.code = code;
        this.statusCode = statusCode;
        this.originalError = originalError;
    }
}
/**
 * Maps an HTTP fetch error (or any unknown thrown value) to a TeamsAdapterError.
 */
export function handleApiError(error) {
    if (error instanceof TeamsAdapterError) {
        return error;
    }
    if (error instanceof Response) {
        return mapHttpResponse(error, undefined);
    }
    if (typeof error === 'object' &&
        error !== null &&
        'status' in error &&
        typeof error['status'] === 'number') {
        const httpErr = error;
        const message = httpErr.body?.error?.message ??
            httpErr.body?.message ??
            httpErr.statusText ??
            'Unknown HTTP error';
        return mapHttpResponse({ status: httpErr.status }, message);
    }
    if (error instanceof Error) {
        return new TeamsAdapterError(error.message, TeamsAdapterErrorCode.API_ERROR, 500, error);
    }
    return new TeamsAdapterError(typeof error === 'string' ? error : 'An unknown error occurred', TeamsAdapterErrorCode.UNKNOWN, 500, error);
}
/**
 * Creates a TeamsAdapterError from a failed HTTP Response object.
 */
export async function handleHttpError(response) {
    let body;
    try {
        body = (await response.json());
    }
    catch {
        // Ignore body parse failure
    }
    const message = body?.error?.message ?? body?.message ?? response.statusText ?? 'HTTP error';
    return mapHttpResponse(response, message);
}
function mapHttpResponse(response, message) {
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
export function isRateLimitError(error) {
    return (error instanceof TeamsAdapterError &&
        error.code === TeamsAdapterErrorCode.RATE_LIMITED);
}
/**
 * Extracts the Retry-After delay in milliseconds from a rate-limit error or
 * a Response headers object. Falls back to the provided default.
 */
export function getRateLimitDelay(headersOrError, defaultDelayMs = 1000) {
    if (headersOrError instanceof Headers) {
        const retryAfter = headersOrError.get('retry-after');
        if (retryAfter) {
            const seconds = Number(retryAfter);
            if (!isNaN(seconds))
                return seconds * 1000;
        }
    }
    return defaultDelayMs;
}
//# sourceMappingURL=error-handler.js.map