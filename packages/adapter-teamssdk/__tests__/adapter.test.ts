import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { TeamsAdapterImpl, createTeamsAdapter, TeamsAdapterError, TeamsAdapterErrorCode } from '../src/index.js';
import type { TeamsAdapterConfig, TeamsActivity, ChatInstance, Message } from '../src/types.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeChatInstance(): ChatInstance {
  return {
    emit: vi.fn(),
    on: vi.fn(),
  };
}

function makeConfig(overrides: Partial<TeamsAdapterConfig> = {}): TeamsAdapterConfig {
  return {
    appId: 'test-app-id',
    appPassword: 'test-password',
    tenantId: 'test-tenant',
    enableLogging: false,
    maxRetries: 0,
    ...overrides,
  };
}

function makeMessageActivity(overrides: Partial<TeamsActivity> = {}): TeamsActivity {
  return {
    type: 'message',
    id: 'activity-1',
    serviceUrl: 'https://smba.trafficmanager.net/apis',
    channelId: 'msteams',
    from: { id: 'user-1', name: 'Test User' },
    conversation: { id: 'conv-1', isGroup: true, tenantId: 'test-tenant' },
    recipient: { id: 'bot-1', name: 'Test Bot' },
    text: 'Hello world',
    timestamp: '2024-01-01T00:00:00Z',
    ...overrides,
  };
}

// Minimal JWT with valid structure (header.payload.signature)
function makeFakeJwt(payload: Record<string, unknown> = {}): string {
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT', kid: 'key1' })).toString('base64url');
  const defaultPayload = {
    aud: 'test-app-id',
    iss: 'https://api.botframework.com',
    exp: Math.floor(Date.now() / 1000) + 3600,
    ...payload,
  };
  const body = Buffer.from(JSON.stringify(defaultPayload)).toString('base64url');
  return `${header}.${body}.fakesig`;
}

// ---------------------------------------------------------------------------
// Tests: Factory & Config
// ---------------------------------------------------------------------------

describe('createTeamsAdapter', () => {
  it('creates an adapter instance', () => {
    const adapter = createTeamsAdapter(makeConfig());
    expect(adapter).toBeInstanceOf(TeamsAdapterImpl);
  });

  it('exposes config', () => {
    const config = makeConfig();
    const adapter = createTeamsAdapter(config);
    expect(adapter.config.appId).toBe('test-app-id');
  });

  it('throws if appId is missing', () => {
    expect(() => createTeamsAdapter({ appId: '' })).toThrow(TeamsAdapterError);
  });
});

// ---------------------------------------------------------------------------
// Tests: initialize
// ---------------------------------------------------------------------------

describe('TeamsAdapterImpl.initialize', () => {
  it('stores the chat instance', async () => {
    const adapter = createTeamsAdapter(makeConfig());
    const chat = makeChatInstance();
    await adapter.initialize(chat);
    expect(adapter.chatInstance).toBe(chat);
  });
});

// ---------------------------------------------------------------------------
// Tests: handleWebhook
// ---------------------------------------------------------------------------

describe('TeamsAdapterImpl.handleWebhook', () => {
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    adapter = createTeamsAdapter(makeConfig());
    await adapter.initialize(makeChatInstance());
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('returns 401 when Authorization header is missing', async () => {
    const result = await adapter.handleWebhook({
      headers: {},
      body: makeMessageActivity(),
    });
    expect(result.status).toBe(401);
  });

  it('returns 401 when token validation fails', async () => {
    // Mock validateToken to return false
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({
        jwks_uri: 'https://example.com/jwks',
      }),
    }));

    const result = await adapter.handleWebhook({
      headers: { authorization: 'Bearer invalid-token' },
      body: makeMessageActivity(),
    });
    // Should return 401 since token validation will fail for invalid JWT
    expect(result.status).toBe(401);
    vi.unstubAllGlobals();
  });

  it('returns 400 when activity body is missing type', async () => {
    // Mock validateToken to pass
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    const result = await adapter.handleWebhook({
      headers: { authorization: 'Bearer valid-token' },
      body: {},
    });
    expect(result.status).toBe(400);
    vi.restoreAllMocks();
  });

  it('returns 200 and processes a message activity', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    const result = await adapter.handleWebhook({
      headers: { authorization: `Bearer ${makeFakeJwt()}` },
      body: makeMessageActivity(),
    });
    expect(result.status).toBe(200);
    vi.restoreAllMocks();
  });

  it('caches serviceUrl from activity', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    const activity = makeMessageActivity({
      serviceUrl: 'https://custom-service-url.example.com',
    });
    await adapter.handleWebhook({
      headers: { authorization: `Bearer ${makeFakeJwt()}` },
      body: activity,
    });
    // Verify service URL was cached by posting a message
    // (just check no error is thrown on thread operations)
    vi.restoreAllMocks();
  });
});

// ---------------------------------------------------------------------------
// Tests: encodeThreadId / decodeThreadId
// ---------------------------------------------------------------------------

describe('Thread ID encoding', () => {
  const adapter = createTeamsAdapter(makeConfig());

  it('round-trips a full context', () => {
    const ctx = {
      serviceUrl: 'https://smba.trafficmanager.net',
      tenantId: 'tenant-1',
      conversationId: 'conv-abc',
      channelId: 'chan-xyz',
      teamId: 'team-123',
    };
    const id = adapter.encodeThreadId(ctx);
    expect(typeof id).toBe('string');
    const decoded = adapter.decodeThreadId(id);
    expect(decoded).toEqual(ctx);
  });

  it('throws on invalid thread ID', () => {
    expect(() => adapter.decodeThreadId('not-valid-base64!!')).toThrow();
  });
});

// ---------------------------------------------------------------------------
// Tests: isDM
// ---------------------------------------------------------------------------

describe('isDM', () => {
  const adapter = createTeamsAdapter(makeConfig());

  it('returns true for DM context', () => {
    const id = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net',
      tenantId: 'tenant',
      conversationId: 'conv-dm',
      isDM: true,
    });
    expect(adapter.isDM(id)).toBe(true);
  });

  it('returns false for channel context', () => {
    const id = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net',
      tenantId: 'tenant',
      conversationId: 'conv-channel',
      channelId: 'chan-1',
      teamId: 'team-1',
    });
    expect(adapter.isDM(id)).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Tests: postMessage
// ---------------------------------------------------------------------------

describe('TeamsAdapterImpl.postMessage', () => {
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    adapter = createTeamsAdapter(makeConfig({ maxRetries: 0 }));
    await adapter.initialize(makeChatInstance());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('POSTs to bot connector and returns SentMessage', async () => {
    // Mock token acquisition
    vi.stubGlobal('fetch', vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({ access_token: 'tok', expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({ id: 'msg-999', timestamp: '2024-01-01T00:00:00Z' }),
      })
    );

    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'tenant',
      conversationId: 'conv-1',
    });

    const sent = await adapter.postMessage(threadId, { text: 'Hello' });
    expect(sent.id).toBe('msg-999');
    expect(sent.timestamp).toBe('2024-01-01T00:00:00Z');
  });

  it('throws TeamsAdapterError on 403', async () => {
    vi.stubGlobal('fetch', vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({ access_token: 'tok', expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: false,
        status: 403,
        statusText: 'Forbidden',
        json: async () => ({ error: { message: 'Forbidden' } }),
      })
    );

    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'tenant',
      conversationId: 'conv-1',
    });

    await expect(adapter.postMessage(threadId, { text: 'Hi' })).rejects.toThrow(TeamsAdapterError);
  });
});

// ---------------------------------------------------------------------------
// Tests: editMessage
// ---------------------------------------------------------------------------

describe('TeamsAdapterImpl.editMessage', () => {
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    adapter = createTeamsAdapter(makeConfig({ maxRetries: 0 }));
    await adapter.initialize(makeChatInstance());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('sends PUT to bot connector', async () => {
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({ access_token: 'tok', expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({}),
      });
    vi.stubGlobal('fetch', mockFetch);

    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'tenant',
      conversationId: 'conv-1',
    });

    await adapter.editMessage(threadId, 'msg-1', { text: 'Updated' });
    const putCall = mockFetch.mock.calls[1] as [string, RequestInit];
    expect(putCall[1].method).toBe('PUT');
  });
});

// ---------------------------------------------------------------------------
// Tests: deleteMessage
// ---------------------------------------------------------------------------

describe('TeamsAdapterImpl.deleteMessage', () => {
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    adapter = createTeamsAdapter(makeConfig({ maxRetries: 0 }));
    await adapter.initialize(makeChatInstance());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('sends DELETE to bot connector', async () => {
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: async () => ({ access_token: 'tok', expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 204,
        json: async () => ({}),
      });
    vi.stubGlobal('fetch', mockFetch);

    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'tenant',
      conversationId: 'conv-1',
    });

    await adapter.deleteMessage(threadId, 'msg-1');
    const deleteCall = mockFetch.mock.calls[1] as [string, RequestInit];
    expect(deleteCall[1].method).toBe('DELETE');
  });
});

// ---------------------------------------------------------------------------
// Tests: parseMessage / renderFormatted
// ---------------------------------------------------------------------------

describe('TeamsAdapterImpl.parseMessage', () => {
  const adapter = createTeamsAdapter(makeConfig());

  it('converts activity to message', () => {
    const activity = makeMessageActivity({ text: 'Hello', id: 'act-1' });
    const msg = adapter.parseMessage(activity);
    expect(msg.text).toBe('Hello');
    expect(msg.id).toBe('act-1');
    expect(msg.userId).toBe('user-1');
  });

  it('strips mentions from text', () => {
    const activity = makeMessageActivity({ text: '<at>TestBot</at> hello' });
    const msg = adapter.parseMessage(activity);
    expect(msg.text).toBe('hello');
  });
});

describe('TeamsAdapterImpl.renderFormatted', () => {
  const adapter = createTeamsAdapter(makeConfig());

  it('renders markdown to HTML', () => {
    const msg: Message = { text: '**bold**' };
    const result = adapter.renderFormatted({
      type: 'markdown',
      content: '**bold**',
    });
    expect(result).toContain('<strong>bold</strong>');
  });

  it('returns empty string for adaptive_card type', () => {
    const result = adapter.renderFormatted({
      type: 'adaptive_card',
      content: { type: 'AdaptiveCard' },
    });
    expect(result).toBe('');
  });
});

// ---------------------------------------------------------------------------
// Tests: fetchThread
// ---------------------------------------------------------------------------

describe('TeamsAdapterImpl.fetchThread', () => {
  it('returns a Thread from threadId', async () => {
    const adapter = createTeamsAdapter(makeConfig());
    const id = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net',
      tenantId: 'tenant',
      conversationId: 'conv-1',
      channelId: 'chan-1',
      teamId: 'team-1',
    });
    const thread = await adapter.fetchThread(id);
    expect(thread.id).toBe(id);
    expect(thread.channelId).toBe('chan-1');
  });
});
