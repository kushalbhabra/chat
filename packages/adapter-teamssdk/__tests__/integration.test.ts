import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { createTeamsAdapter } from '../src/index.js';
import type { TeamsAdapterImpl } from '../src/index.js';
import type { ChatInstance, TeamsActivity } from '../src/types.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeChat(): ChatInstance & { emitted: Array<[string, unknown]> } {
  const emitted: Array<[string, unknown]> = [];
  return {
    emitted,
    emit: vi.fn((event: string, data: unknown) => {
      emitted.push([event, data]);
    }),
    on: vi.fn(),
  };
}

function makeAdapter(): TeamsAdapterImpl {
  return createTeamsAdapter({
    appId: 'integration-app-id',
    appPassword: 'integration-password',
    tenantId: 'integration-tenant',
    enableLogging: false,
    maxRetries: 0,
  });
}

function makeTokenFetch(): typeof fetch {
  return vi.fn().mockResolvedValue({
    ok: true,
    status: 200,
    json: async () => ({ access_token: 'integration-token', expires_in: 3600 }),
  }) as unknown as typeof fetch;
}

// ---------------------------------------------------------------------------
// Integration: full message flow
// ---------------------------------------------------------------------------

describe('Full message flow: receive webhook → emit event → post reply', () => {
  let adapter: TeamsAdapterImpl;
  let chat: ReturnType<typeof makeChat>;

  beforeEach(async () => {
    chat = makeChat();
    adapter = makeAdapter();
    await adapter.initialize(chat);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('processes a message activity and emits the message event', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    const activity: TeamsActivity = {
      type: 'message',
      id: 'integ-act-1',
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      channelId: 'msteams',
      from: { id: 'user-integ', name: 'Integration User' },
      conversation: { id: 'conv-integ', isGroup: true, tenantId: 'integration-tenant' },
      recipient: { id: 'integration-app-id', name: 'IntegBot' },
      text: 'Integration test message',
      timestamp: '2024-06-01T12:00:00Z',
    };

    const result = await adapter.handleWebhook({
      headers: { authorization: 'Bearer valid-token' },
      body: activity,
    });

    expect(result.status).toBe(200);
    expect(chat.emitted.some(([e]) => e === 'message')).toBe(true);
    const [, data] = chat.emitted.find(([e]) => e === 'message')!;
    expect((data as Record<string, unknown>)['message']).toMatchObject({
      text: 'Integration test message',
      userId: 'user-integ',
    });
  });

  it('posts a reply message after receiving a webhook', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    // First call: validate token via JWKS (covered by spy above)
    // Subsequent calls: token acquisition + post message
    vi.stubGlobal('fetch', vi.fn()
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({ access_token: 'tok', expires_in: 3600 }) })
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({ id: 'reply-1', timestamp: '2024-06-01T12:00:01Z' }) })
    );

    const activity: TeamsActivity = {
      type: 'message',
      id: 'integ-act-2',
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      channelId: 'msteams',
      from: { id: 'user-integ', name: 'Integration User' },
      conversation: { id: 'conv-reply', isGroup: true, tenantId: 'integration-tenant' },
      recipient: { id: 'integration-app-id', name: 'IntegBot' },
      text: 'Please reply',
      timestamp: '2024-06-01T12:00:00Z',
    };

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer valid-token' },
      body: activity,
    });

    // Now post a reply using the conversation ID we just received
    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'integration-tenant',
      conversationId: 'conv-reply',
    });

    const sent = await adapter.postMessage(threadId, { text: 'Here is my reply!' });
    expect(sent.id).toBe('reply-1');
    expect(sent.timestamp).toBe('2024-06-01T12:00:01Z');
  });
});

// ---------------------------------------------------------------------------
// Integration: DM flow
// ---------------------------------------------------------------------------

describe('DM flow', () => {
  let adapter: TeamsAdapterImpl;
  let chat: ReturnType<typeof makeChat>;

  beforeEach(async () => {
    chat = makeChat();
    adapter = makeAdapter();
    await adapter.initialize(chat);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('routes a personal-scope message to dm_received', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    const activity: TeamsActivity = {
      type: 'message',
      id: 'dm-act-1',
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      channelId: 'msteams',
      from: { id: 'user-dm', name: 'DM User' },
      conversation: { id: 'dm-conv', conversationType: 'personal', isGroup: false, tenantId: 'integration-tenant' },
      recipient: { id: 'integration-app-id', name: 'IntegBot' },
      text: 'This is a DM',
      timestamp: '2024-06-01T12:00:00Z',
    };

    const result = await adapter.handleWebhook({
      headers: { authorization: 'Bearer valid-token' },
      body: activity,
    });

    expect(result.status).toBe(200);
    expect(chat.emitted.some(([e]) => e === 'dm_received')).toBe(true);
    expect(chat.emitted.some(([e]) => e === 'message')).toBe(false);
  });

  it('openDM creates a one-on-one chat and returns encoded thread ID', async () => {
    vi.stubGlobal('fetch', vi.fn()
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({ access_token: 'tok', expires_in: 3600 }) })
      .mockResolvedValueOnce({ ok: true, status: 201, json: async () => ({ id: 'chat-dm-created' }) })
    );

    const threadId = await adapter.openDM('user-target');
    expect(typeof threadId).toBe('string');

    const decoded = adapter.decodeThreadId(threadId);
    expect(decoded.conversationId).toBe('chat-dm-created');
    expect(decoded.isDM).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Integration: reaction flow
// ---------------------------------------------------------------------------

describe('Reaction flow', () => {
  let adapter: TeamsAdapterImpl;
  let chat: ReturnType<typeof makeChat>;

  beforeEach(async () => {
    chat = makeChat();
    adapter = makeAdapter();
    await adapter.initialize(chat);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('emits reaction_added when messageReaction activity arrives with reactionsAdded', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    const activity: TeamsActivity = {
      type: 'messageReaction',
      id: 'react-act-1',
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      channelId: 'msteams',
      from: { id: 'user-reactor', name: 'Reactor' },
      conversation: { id: 'react-conv', isGroup: true, tenantId: 'integration-tenant' },
      recipient: { id: 'integration-app-id', name: 'IntegBot' },
      reactionsAdded: [{ type: 'like' }],
      replyToId: 'msg-being-reacted-to',
    };

    const result = await adapter.handleWebhook({
      headers: { authorization: 'Bearer valid-token' },
      body: activity,
    });

    expect(result.status).toBe(200);
    expect(chat.emitted.some(([e]) => e === 'reaction_added')).toBe(true);
    const [, data] = chat.emitted.find(([e]) => e === 'reaction_added')!;
    expect((data as Record<string, unknown>)['reactions']).toEqual([{ type: 'like' }]);
    expect((data as Record<string, unknown>)['messageId']).toBe('msg-being-reacted-to');
  });

  it('calls Graph API addReaction and does not throw', async () => {
    vi.stubGlobal('fetch', vi.fn()
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({ access_token: 'tok', expires_in: 3600 }) })
      .mockResolvedValueOnce({ ok: true, status: 204, json: async () => ({}) })
    );

    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'integration-tenant',
      conversationId: 'react-conv',
      teamId: 'team-1',
      channelId: 'chan-1',
    });

    await expect(adapter.addReaction(threadId, 'msg-1', 'like')).resolves.not.toThrow();
  });
});

// ---------------------------------------------------------------------------
// Integration: startTyping
// ---------------------------------------------------------------------------

describe('startTyping', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('POSTs a typing activity', async () => {
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({ access_token: 'tok', expires_in: 3600 }) })
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({}) });
    vi.stubGlobal('fetch', mockFetch);

    const adapter = makeAdapter();
    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'tenant',
      conversationId: 'conv-typing',
    });

    await adapter.startTyping(threadId);
    const [, init] = mockFetch.mock.calls[1] as [string, RequestInit];
    const body = JSON.parse(init.body as string) as { type: string };
    expect(body.type).toBe('typing');
  });
});

// ---------------------------------------------------------------------------
// Integration: retry on transient error
// ---------------------------------------------------------------------------

describe('Retry logic', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('retries on 500 and succeeds on second attempt', async () => {
    const adapter = createTeamsAdapter({
      appId: 'retry-app',
      appPassword: 'retry-pass',
      maxRetries: 2,
      retryDelayMs: 1,
      enableLogging: false,
    });

    const mockFetch = vi.fn()
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({ access_token: 'tok', expires_in: 3600 }) })
      .mockResolvedValueOnce({ ok: false, status: 500, statusText: 'Internal Server Error', json: async () => ({}) })
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({ id: 'retry-success', timestamp: 'ts' }) });
    vi.stubGlobal('fetch', mockFetch);

    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'tenant',
      conversationId: 'conv-retry',
    });

    const sent = await adapter.postMessage(threadId, { text: 'Will retry' });
    expect(sent.id).toBe('retry-success');
    // Token call + 1 failed + 1 success = 3 calls
    expect(mockFetch).toHaveBeenCalledTimes(3);
  });

  it('throws after exhausting retries', async () => {
    const adapter = createTeamsAdapter({
      appId: 'retry-app',
      appPassword: 'retry-pass',
      maxRetries: 1,
      retryDelayMs: 1,
      enableLogging: false,
    });

    const mockFetch = vi.fn()
      .mockResolvedValueOnce({ ok: true, status: 200, json: async () => ({ access_token: 'tok', expires_in: 3600 }) })
      .mockResolvedValue({ ok: false, status: 500, statusText: 'Server Error', json: async () => ({}) });
    vi.stubGlobal('fetch', mockFetch);

    const threadId = adapter.encodeThreadId({
      serviceUrl: 'https://smba.trafficmanager.net/apis',
      tenantId: 'tenant',
      conversationId: 'conv-exhaust',
    });

    await expect(adapter.postMessage(threadId, { text: 'fail' })).rejects.toThrow();
  });
});
