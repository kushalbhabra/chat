import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  encodeThreadId,
  decodeThreadId,
  channelIdFromThreadId,
  isDM,
} from '../src/utils/thread-utils.js';
import {
  activityToMessage,
  messageToActivity,
  renderMarkdown,
  adaptiveCardToAttachment,
} from '../src/utils/format-converter.js';
import {
  TeamsAdapterError,
  TeamsAdapterErrorCode,
  handleApiError,
  isRateLimitError,
} from '../src/utils/error-handler.js';
import type { TeamsActivity, TeamsContext, FormattedContent } from '../src/types.js';

// ---------------------------------------------------------------------------
// Thread Utils
// ---------------------------------------------------------------------------

describe('encodeThreadId / decodeThreadId', () => {
  it('round-trips a minimal context', () => {
    const ctx: TeamsContext = {
      serviceUrl: 'https://smba.trafficmanager.net',
      tenantId: 'tenant-1',
      conversationId: 'conv-abc',
    };
    const id = encodeThreadId(ctx);
    expect(typeof id).toBe('string');
    expect(id.length).toBeGreaterThan(0);
    const decoded = decodeThreadId(id);
    expect(decoded).toEqual(ctx);
  });

  it('round-trips a full context with teamId and channelId', () => {
    const ctx: TeamsContext = {
      serviceUrl: 'https://smba.trafficmanager.net',
      tenantId: 'tenant-xyz',
      conversationId: 'conv-999',
      channelId: 'chan-123',
      teamId: 'team-456',
      isDM: false,
    };
    const id = encodeThreadId(ctx);
    const decoded = decodeThreadId(id);
    expect(decoded).toEqual(ctx);
  });

  it('throws on non-base64 input', () => {
    expect(() => decodeThreadId('not!valid')).toThrow();
  });

  it('throws on base64 that is not valid JSON', () => {
    const badB64 = Buffer.from('not json').toString('base64url');
    expect(() => decodeThreadId(badB64)).toThrow();
  });

  it('throws when required fields are missing', () => {
    const badB64 = Buffer.from(JSON.stringify({ foo: 'bar' })).toString('base64url');
    expect(() => decodeThreadId(badB64)).toThrow('missing required fields');
  });
});

describe('channelIdFromThreadId', () => {
  it('returns channelId when present', () => {
    const id = encodeThreadId({
      serviceUrl: '',
      tenantId: 't',
      conversationId: 'conv',
      channelId: 'chan-xyz',
    });
    expect(channelIdFromThreadId(id)).toBe('chan-xyz');
  });

  it('falls back to conversationId when channelId absent', () => {
    const id = encodeThreadId({
      serviceUrl: '',
      tenantId: 't',
      conversationId: 'conv-fallback',
    });
    expect(channelIdFromThreadId(id)).toBe('conv-fallback');
  });
});

describe('isDM', () => {
  it('returns true when isDM flag is set', () => {
    const id = encodeThreadId({
      serviceUrl: '',
      tenantId: 't',
      conversationId: 'c',
      isDM: true,
    });
    expect(isDM(id)).toBe(true);
  });

  it('returns false when isDM flag is false', () => {
    const id = encodeThreadId({
      serviceUrl: '',
      tenantId: 't',
      conversationId: 'c',
      isDM: false,
    });
    expect(isDM(id)).toBe(false);
  });

  it('infers DM when no channelId or teamId', () => {
    const id = encodeThreadId({
      serviceUrl: '',
      tenantId: 't',
      conversationId: 'c',
    });
    expect(isDM(id)).toBe(true);
  });

  it('infers not-DM when channelId is present', () => {
    const id = encodeThreadId({
      serviceUrl: '',
      tenantId: 't',
      conversationId: 'c',
      channelId: 'chan-1',
    });
    expect(isDM(id)).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Format Converter
// ---------------------------------------------------------------------------

function makeActivity(overrides: Partial<TeamsActivity> = {}): TeamsActivity {
  return {
    type: 'message',
    id: 'act-1',
    serviceUrl: 'https://smba.example.com',
    channelId: 'msteams',
    from: { id: 'user-1', name: 'Alice', aadObjectId: 'aad-1' },
    conversation: { id: 'conv-1', isGroup: true, tenantId: 'tenant-1' },
    recipient: { id: 'bot-1', name: 'TestBot' },
    text: 'Hello',
    timestamp: '2024-01-01T00:00:00Z',
    ...overrides,
  };
}

describe('activityToMessage', () => {
  it('maps basic fields correctly', () => {
    const activity = makeActivity();
    const msg = activityToMessage(activity);
    expect(msg.id).toBe('act-1');
    expect(msg.text).toBe('Hello');
    expect(msg.userId).toBe('user-1');
    expect(msg.userName).toBe('Alice');
    expect(msg.timestamp).toBe('2024-01-01T00:00:00Z');
    expect(msg.threadId).toBe('conv-1');
  });

  it('strips @mention tags from text', () => {
    const activity = makeActivity({ text: '<at>TestBot</at> what time is it?' });
    const msg = activityToMessage(activity);
    expect(msg.text).toBe('what time is it?');
  });

  it('handles empty text gracefully', () => {
    const activity = makeActivity({ text: undefined });
    const msg = activityToMessage(activity);
    expect(msg.text).toBe('');
  });

  it('converts adaptive card attachments', () => {
    const activity = makeActivity({
      attachments: [
        {
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: { type: 'AdaptiveCard' },
        },
      ],
    });
    const msg = activityToMessage(activity);
    expect(msg.attachments).toHaveLength(1);
    expect(msg.attachments![0]!.type).toBe('card');
  });

  it('converts image attachments', () => {
    const activity = makeActivity({
      attachments: [{ contentType: 'image/png', contentUrl: 'https://example.com/img.png', name: 'img.png' }],
    });
    const msg = activityToMessage(activity);
    expect(msg.attachments![0]!.type).toBe('image');
    expect(msg.attachments![0]!.url).toBe('https://example.com/img.png');
  });
});

describe('messageToActivity', () => {
  it('creates a message activity from text', () => {
    const activity = messageToActivity({ text: 'Hello Teams' });
    expect(activity.type).toBe('message');
    expect(activity.text).toBe('Hello Teams');
    expect(activity.textFormat).toBe('markdown');
  });

  it('renders markdown content', () => {
    const activity = messageToActivity({
      formattedContent: { type: 'markdown', content: '**Bold**' },
    });
    expect(activity.text).toContain('<strong>Bold</strong>');
  });

  it('preserves replyToId', () => {
    const activity = messageToActivity({ text: 'reply', replyToId: 'parent-1' });
    expect(activity.replyToId).toBe('parent-1');
  });

  it('converts card attachments back to Teams format', () => {
    const activity = messageToActivity({
      text: 'with card',
      attachments: [{ type: 'card', content: { type: 'AdaptiveCard' }, contentType: 'application/vnd.microsoft.card.adaptive' }],
    });
    expect(activity.attachments).toHaveLength(1);
    expect(activity.attachments![0]!.contentType).toBe('application/vnd.microsoft.card.adaptive');
  });
});

describe('renderMarkdown', () => {
  const cases: Array<[string, FormattedContent, string]> = [
    ['bold', { type: 'markdown', content: '**bold text**' }, '<strong>bold text</strong>'],
    ['italic', { type: 'markdown', content: '*italic text*' }, '<em>italic text</em>'],
    ['strikethrough', { type: 'markdown', content: '~~struck~~' }, '<del>struck</del>'],
    ['inline code', { type: 'markdown', content: '`code`' }, '<code>code</code>'],
    ['link', { type: 'markdown', content: '[Click](https://example.com)' }, '<a href="https://example.com">Click</a>'],
    ['HTML passthrough', { type: 'html', content: '<b>raw html</b>' }, '<b>raw html</b>'],
    ['plain text escapes HTML', { type: 'text', content: '<script>' }, '&lt;script&gt;'],
    ['adaptive card returns empty', { type: 'adaptive_card', content: {} }, ''],
  ];

  for (const [name, input, expected] of cases) {
    it(`renders ${name}`, () => {
      const result = renderMarkdown(input);
      expect(result).toContain(expected);
    });
  }

  it('renders blocks type', () => {
    const result = renderMarkdown({
      type: 'blocks',
      content: '',
      blocks: [{ type: 'section', text: 'Hello blocks' }],
    });
    expect(result).toContain('Hello blocks');
  });
});

describe('adaptiveCardToAttachment', () => {
  it('wraps card in AdaptiveCard envelope', () => {
    const att = adaptiveCardToAttachment({ body: [{ type: 'TextBlock', text: 'Hi' }] });
    expect(att.contentType).toBe('application/vnd.microsoft.card.adaptive');
    expect((att.content as Record<string, unknown>)['type']).toBe('AdaptiveCard');
    expect((att.content as Record<string, unknown>)['version']).toBe('1.4');
  });

  it('merges custom fields', () => {
    const att = adaptiveCardToAttachment({ version: '1.5', actions: [] });
    expect((att.content as Record<string, unknown>)['version']).toBe('1.5');
  });
});

// ---------------------------------------------------------------------------
// Error Handler
// ---------------------------------------------------------------------------

describe('TeamsAdapterError', () => {
  it('has correct properties', () => {
    const err = new TeamsAdapterError('test', TeamsAdapterErrorCode.NOT_FOUND, 404);
    expect(err.message).toBe('test');
    expect(err.code).toBe(TeamsAdapterErrorCode.NOT_FOUND);
    expect(err.statusCode).toBe(404);
    expect(err.name).toBe('TeamsAdapterError');
  });
});

describe('handleApiError', () => {
  it('passes through a TeamsAdapterError unchanged', () => {
    const original = new TeamsAdapterError('original', TeamsAdapterErrorCode.FORBIDDEN, 403);
    const result = handleApiError(original);
    expect(result).toBe(original);
  });

  it('wraps a generic Error', () => {
    const err = handleApiError(new Error('generic'));
    expect(err).toBeInstanceOf(TeamsAdapterError);
    expect(err.code).toBe(TeamsAdapterErrorCode.API_ERROR);
    expect(err.message).toBe('generic');
  });

  it('wraps a string', () => {
    const err = handleApiError('oops');
    expect(err.message).toBe('oops');
  });

  it('wraps an object with status 401', () => {
    const err = handleApiError({ status: 401, statusText: 'Unauthorized' });
    expect(err.code).toBe(TeamsAdapterErrorCode.UNAUTHORIZED);
    expect(err.statusCode).toBe(401);
  });

  it('wraps an object with status 429', () => {
    const err = handleApiError({ status: 429, statusText: 'Too Many Requests' });
    expect(err.code).toBe(TeamsAdapterErrorCode.RATE_LIMITED);
  });

  it('wraps an object with status 404', () => {
    const err = handleApiError({ status: 404, statusText: 'Not Found' });
    expect(err.code).toBe(TeamsAdapterErrorCode.NOT_FOUND);
  });
});

describe('isRateLimitError', () => {
  it('returns true for RATE_LIMITED errors', () => {
    const err = new TeamsAdapterError('rate limit', TeamsAdapterErrorCode.RATE_LIMITED, 429);
    expect(isRateLimitError(err)).toBe(true);
  });

  it('returns false for other errors', () => {
    const err = new TeamsAdapterError('not found', TeamsAdapterErrorCode.NOT_FOUND, 404);
    expect(isRateLimitError(err)).toBe(false);
  });

  it('returns false for non-TeamsAdapterError', () => {
    expect(isRateLimitError(new Error('plain'))).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Auth validator (mocked)
// ---------------------------------------------------------------------------

describe('validateToken', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('returns false for a malformed token (no dots)', async () => {
    const { validateToken } = await import('../src/utils/auth-validator.js');
    const result = await validateToken('not-a-jwt', 'app-id');
    expect(result).toBe(false);
  });

  it('returns false when token is expired', async () => {
    const { validateToken } = await import('../src/utils/auth-validator.js');
    const header = Buffer.from(JSON.stringify({ alg: 'RS256', kid: 'k1' })).toString('base64url');
    const payload = Buffer.from(
      JSON.stringify({ aud: 'app-id', iss: 'https://api.botframework.com', exp: 1 })
    ).toString('base64url');
    const token = `${header}.${payload}.sig`;
    const result = await validateToken(token, 'app-id');
    expect(result).toBe(false);
  });

  it('returns false when audience does not match', async () => {
    const { validateToken } = await import('../src/utils/auth-validator.js');
    const header = Buffer.from(JSON.stringify({ alg: 'RS256', kid: 'k1' })).toString('base64url');
    const payload = Buffer.from(
      JSON.stringify({
        aud: 'wrong-app',
        iss: 'https://api.botframework.com',
        exp: Math.floor(Date.now() / 1000) + 3600,
      })
    ).toString('base64url');
    const token = `${header}.${payload}.sig`;
    const result = await validateToken(token, 'app-id');
    expect(result).toBe(false);
  });

  it('returns false when issuer is invalid', async () => {
    const { validateToken } = await import('../src/utils/auth-validator.js');
    const header = Buffer.from(JSON.stringify({ alg: 'RS256', kid: 'k1' })).toString('base64url');
    const payload = Buffer.from(
      JSON.stringify({
        aud: 'app-id',
        iss: 'https://malicious.example.com',
        exp: Math.floor(Date.now() / 1000) + 3600,
      })
    ).toString('base64url');
    const token = `${header}.${payload}.sig`;
    const result = await validateToken(token, 'app-id');
    expect(result).toBe(false);
  });

  it('returns false when JWKS fetch fails', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('network error')));
    const { validateToken } = await import('../src/utils/auth-validator.js');
    const header = Buffer.from(JSON.stringify({ alg: 'RS256', kid: 'k1' })).toString('base64url');
    const payload = Buffer.from(
      JSON.stringify({
        aud: 'app-id',
        iss: 'https://api.botframework.com',
        exp: Math.floor(Date.now() / 1000) + 3600,
      })
    ).toString('base64url');
    const token = `${header}.${payload}.sig`;
    const result = await validateToken(token, 'app-id');
    expect(result).toBe(false);
  });

  it('returns false when kid not found in JWKS', async () => {
    vi.stubGlobal('fetch', vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: async () => ({ jwks_uri: 'https://example.com/jwks' }),
      })
      .mockResolvedValueOnce({
        ok: true,
        json: async () => ({ keys: [{ kid: 'other-key', n: 'abc', e: 'AQAB' }] }),
      })
    );
    const { validateToken } = await import('../src/utils/auth-validator.js');
    const header = Buffer.from(JSON.stringify({ alg: 'RS256', kid: 'missing-key' })).toString('base64url');
    const payload = Buffer.from(
      JSON.stringify({
        aud: 'app-id',
        iss: 'https://api.botframework.com',
        exp: Math.floor(Date.now() / 1000) + 3600,
      })
    ).toString('base64url');
    const token = `${header}.${payload}.sig`;
    const result = await validateToken(token, 'app-id');
    expect(result).toBe(false);
  });
});
