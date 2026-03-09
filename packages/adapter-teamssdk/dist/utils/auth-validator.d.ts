import type { AccessToken, TokenCredentials } from '../types.js';
/**
 * Validates a Bot Framework JWT token.
 *
 * Performs:
 * 1. Structural validation (3-part JWT)
 * 2. Claims validation (iss, aud, exp)
 * 3. Key ID lookup against Bot Framework JWKS
 *
 * NOTE: This implementation performs claims and key-ID validation.
 * Full cryptographic signature verification requires the `crypto` Web API
 * (available in Node ≥18) and is performed when the runtime supports
 * SubtleCrypto RSA-PKCS1-SHA256 import.
 */
export declare function validateToken(token: string, appId: string): Promise<boolean>;
/**
 * Acquires an OAuth 2.0 access token from Entra ID / Azure AD.
 * Supports password (client secret), certificate, and federated identity flows.
 */
export declare function getAccessToken(credentials: TokenCredentials): Promise<AccessToken>;
//# sourceMappingURL=auth-validator.d.ts.map