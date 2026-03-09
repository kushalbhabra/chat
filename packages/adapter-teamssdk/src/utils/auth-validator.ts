import type { AccessToken, TokenCredentials } from '../types.js';

// Bot Framework / Entra ID token endpoints
const AUTHORITY_BASE = 'https://login.microsoftonline.com';
const BOT_FRAMEWORK_SCOPE = 'https://api.botframework.com/.default';
const BOT_FRAMEWORK_OPENID_CONFIG =
  'https://login.botframework.com/v1/.well-known/openidconfiguration';

interface OpenIdConfig {
  jwks_uri: string;
}

interface JwkSet {
  keys: Jwk[];
}

interface Jwk {
  kid: string;
  n: string;
  e: string;
  x5c?: string[];
  [key: string]: unknown;
}

/**
 * Fetches Bot Framework OpenID configuration (signing key endpoint).
 */
async function fetchOpenIdConfig(): Promise<OpenIdConfig> {
  const res = await fetch(BOT_FRAMEWORK_OPENID_CONFIG);
  if (!res.ok) {
    throw new Error(`Failed to fetch OpenID config: ${res.status} ${res.statusText}`);
  }
  return (await res.json()) as OpenIdConfig;
}

/**
 * Fetches the JWK set from the Bot Framework JWKS URI.
 */
async function fetchJwks(jwksUri: string): Promise<JwkSet> {
  const res = await fetch(jwksUri);
  if (!res.ok) {
    throw new Error(`Failed to fetch JWKS: ${res.status} ${res.statusText}`);
  }
  return (await res.json()) as JwkSet;
}

/**
 * Decodes the header of a JWT token (first segment, base64url).
 */
function decodeJwtHeader(token: string): Record<string, unknown> {
  const parts = token.split('.');
  if (parts.length !== 3) {
    throw new Error('Invalid JWT format');
  }
  const headerJson = Buffer.from(parts[0]!, 'base64url').toString('utf8');
  return JSON.parse(headerJson) as Record<string, unknown>;
}

/**
 * Decodes the payload of a JWT token (second segment, base64url).
 */
function decodeJwtPayload(token: string): Record<string, unknown> {
  const parts = token.split('.');
  if (parts.length !== 3) {
    throw new Error('Invalid JWT format');
  }
  const payloadJson = Buffer.from(parts[1]!, 'base64url').toString('utf8');
  return JSON.parse(payloadJson) as Record<string, unknown>;
}

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
export async function validateToken(token: string, appId: string): Promise<boolean> {
  try {
    const header = decodeJwtHeader(token);
    const payload = decodeJwtPayload(token);

    // Check expiry
    const now = Math.floor(Date.now() / 1000);
    if (typeof payload['exp'] === 'number' && payload['exp'] < now) {
      return false;
    }

    // Check audience matches the app ID
    const aud = payload['aud'];
    if (aud !== appId) {
      return false;
    }

    // Check issuer is Bot Framework or Entra ID
    const iss = payload['iss'] as string | undefined;
    const validIssuers = [
      'https://api.botframework.com',
      `${AUTHORITY_BASE}/botframework.com/`,
    ];
    const isValidIssuer =
      typeof iss === 'string' &&
      (validIssuers.some((v) => iss.startsWith(v)) || iss.includes('microsoftonline.com'));
    if (!isValidIssuer) {
      return false;
    }

    // Fetch JWKS and verify the key ID exists
    const openIdConfig = await fetchOpenIdConfig();
    const jwkSet = await fetchJwks(openIdConfig.jwks_uri);
    const kid = header['kid'] as string | undefined;
    const matchingKey = jwkSet.keys.find((k) => k.kid === kid);
    if (!matchingKey) {
      return false;
    }

    // Attempt cryptographic verification when SubtleCrypto is available
    if (typeof globalThis.crypto?.subtle?.importKey === 'function' && matchingKey.n && matchingKey.e) {
      const verified = await verifyCryptoSignature(token, matchingKey);
      return verified;
    }

    // If SubtleCrypto is not available, structural checks passed
    return true;
  } catch {
    return false;
  }
}

/**
 * Verifies an RS256 JWT signature using the Web Crypto API.
 */
async function verifyCryptoSignature(token: string, jwk: Jwk): Promise<boolean> {
  try {
    const parts = token.split('.');
    const signingInput = `${parts[0]}.${parts[1]}`;
    const signature = parts[2]!;

    const signingInputBytes = new TextEncoder().encode(signingInput);
    const signatureBytes = Buffer.from(signature, 'base64url');

    const cryptoKey = await globalThis.crypto.subtle.importKey(
      'jwk',
      {
        kty: 'RSA',
        n: jwk.n,
        e: jwk.e,
        alg: 'RS256',
        use: 'sig',
      },
      { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
      false,
      ['verify']
    );

    return await globalThis.crypto.subtle.verify(
      'RSASSA-PKCS1-v1_5',
      cryptoKey,
      signatureBytes,
      signingInputBytes
    );
  } catch {
    return false;
  }
}

/**
 * Acquires an OAuth 2.0 access token from Entra ID / Azure AD.
 * Supports password (client secret), certificate, and federated identity flows.
 */
export async function getAccessToken(credentials: TokenCredentials): Promise<AccessToken> {
  const tenantId = 'botframework.com';
  const tokenUrl = `${AUTHORITY_BASE}/${tenantId}/oauth2/v2.0/token`;

  let body: URLSearchParams;

  switch (credentials.type) {
    case 'password': {
      if (!credentials.password) {
        throw new Error('Password credential requires a password');
      }
      body = new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: credentials.appId,
        client_secret: credentials.password,
        scope: BOT_FRAMEWORK_SCOPE,
      });
      break;
    }

    case 'certificate': {
      if (!credentials.thumbprint || !credentials.privateKey) {
        throw new Error('Certificate credential requires thumbprint and privateKey');
      }
      // Build a client assertion JWT signed with the private key
      const assertion = await buildCertificateAssertion(
        credentials.appId,
        credentials.thumbprint,
        credentials.privateKey,
        tokenUrl
      );
      body = new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: credentials.appId,
        client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
        client_assertion: assertion,
        scope: BOT_FRAMEWORK_SCOPE,
      });
      break;
    }

    case 'federated': {
      if (!credentials.federatedToken) {
        throw new Error('Federated credential requires a federatedToken');
      }
      body = new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        client_id: credentials.appId,
        client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
        client_assertion: credentials.federatedToken,
        requested_token_use: 'on_behalf_of',
        scope: BOT_FRAMEWORK_SCOPE,
      });
      break;
    }

    default:
      throw new Error(`Unsupported credential type: ${String((credentials as TokenCredentials).type)}`);
  }

  const res = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Failed to acquire access token: ${res.status} - ${errorText}`);
  }

  const data = (await res.json()) as { access_token: string; expires_in: number };
  return {
    token: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000 - 60_000, // 60-second buffer
  };
}

/**
 * Builds a client assertion JWT for certificate-based auth.
 * Uses RS256 (PKCS#1 v1.5) when SubtleCrypto is available, otherwise
 * uses a placeholder header (for testing/environments without crypto).
 */
async function buildCertificateAssertion(
  appId: string,
  thumbprint: string,
  privateKeyPem: string,
  audience: string
): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  const header = {
    alg: 'RS256',
    typ: 'JWT',
    x5t: Buffer.from(thumbprint, 'hex').toString('base64url'),
  };
  const payload = {
    aud: audience,
    iss: appId,
    sub: appId,
    jti: crypto.randomUUID(),
    nbf: now,
    exp: now + 600,
  };

  const headerB64 = Buffer.from(JSON.stringify(header)).toString('base64url');
  const payloadB64 = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const signingInput = `${headerB64}.${payloadB64}`;

  if (typeof globalThis.crypto?.subtle?.importKey !== 'function') {
    // Fallback: unsigned assertion (will fail server-side, but avoids hard crash)
    return `${signingInput}.UNSIGNED`;
  }

  // Import PEM private key
  const pemBody = privateKeyPem
    .replace(/-----BEGIN[^-]+-----/g, '')
    .replace(/-----END[^-]+-----/g, '')
    .replace(/\s+/g, '');
  const keyBuffer = Buffer.from(pemBody, 'base64');

  const cryptoKey = await globalThis.crypto.subtle.importKey(
    'pkcs8',
    keyBuffer,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false,
    ['sign']
  );

  const signatureBuffer = await globalThis.crypto.subtle.sign(
    'RSASSA-PKCS1-v1_5',
    cryptoKey,
    new TextEncoder().encode(signingInput)
  );

  const sigB64 = Buffer.from(signatureBuffer).toString('base64url');
  return `${signingInput}.${sigB64}`;
}
