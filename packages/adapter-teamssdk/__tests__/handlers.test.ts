import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { createTeamsAdapter } from '../src/index.js';
import type { TeamsAdapterImpl } from '../src/index.js';
import type { TeamsActivity, ChatInstance } from '../src/types.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeChatInstance(): ChatInstance {
  return {
    emit: vi.fn(),
    on: vi.fn(),
  };
}

function makeAdapter(chat?: ChatInstance): TeamsAdapterImpl {
  const adapter = createTeamsAdapter({
    appId: 'test-app-id',
    appPassword: 'test-password',
    enableLogging: false,
  });
  if (chat) {
    adapter.chatInstance = chat;
  }
  return adapter;
}

function makeActivity(overrides: Partial<TeamsActivity> = {}): TeamsActivity {
  return {
    type: 'message',
    id: 'act-1',
    serviceUrl: 'https://smba.example.com',
    channelId: 'msteams',
    from: { id: 'user-1', name: 'User One' },
    conversation: { id: 'conv-1', isGroup: true, tenantId: 'tenant-1' },
    recipient: { id: 'bot-1', name: 'Test Bot' },
    text: 'Hello',
    timestamp: '2024-01-01T00:00:00Z',
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Message handlers
// ---------------------------------------------------------------------------

describe('Message handlers', () => {
  let chat: ChatInstance;
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    chat = makeChatInstance();
    adapter = makeAdapter(chat);
    await adapter.initialize(chat);
  });

  it('emits "message" for a group message', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({ text: 'Hello group' }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'message',
      expect.objectContaining({ message: expect.objectContaining({ text: 'Hello group' }) })
    );
    vi.restoreAllMocks();
  });

  it('emits "dm_received" for a personal conversation', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({
        conversation: { id: 'conv-dm', conversationType: 'personal', isGroup: false },
      }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'dm_received',
      expect.objectContaining({ message: expect.objectContaining({ userId: 'user-1' }) })
    );
    vi.restoreAllMocks();
  });

  it('emits "mention" when bot is mentioned', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({
        entities: [
          { type: 'mention', mentioned: { id: 'bot-1', name: 'Test Bot' }, text: '<at>Test Bot</at>' },
        ],
      }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'mention',
      expect.objectContaining({ message: expect.objectContaining({ userId: 'user-1' }) })
    );
    vi.restoreAllMocks();
  });

  it('emits "thread_reply" when replyToId is set', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({ replyToId: 'parent-msg-1' }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'thread_reply',
      expect.objectContaining({ replyToId: 'parent-msg-1' })
    );
    vi.restoreAllMocks();
  });
});

// ---------------------------------------------------------------------------
// Reaction handlers
// ---------------------------------------------------------------------------

describe('Reaction handlers', () => {
  let chat: ChatInstance;
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    chat = makeChatInstance();
    adapter = makeAdapter(chat);
    await adapter.initialize(chat);
  });

  afterEach(() => vi.restoreAllMocks());

  it('emits "reaction_added"', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({ type: 'messageReaction', reactionsAdded: [{ type: 'like' }] }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'reaction_added',
      expect.objectContaining({ reactions: [{ type: 'like' }] })
    );
  });

  it('emits "reaction_removed"', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({ type: 'messageReaction', reactionsRemoved: [{ type: 'heart' }] }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'reaction_removed',
      expect.objectContaining({ reactions: [{ type: 'heart' }] })
    );
  });
});

// ---------------------------------------------------------------------------
// Card action handlers
// ---------------------------------------------------------------------------

describe('Card action handlers', () => {
  let chat: ChatInstance;
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    chat = makeChatInstance();
    adapter = makeAdapter(chat);
    await adapter.initialize(chat);
  });

  afterEach(() => vi.restoreAllMocks());

  it('emits "card_action" for task/submit invoke', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({
        type: 'invoke',
        name: 'task/submit',
        value: { action: 'confirm', data: { key: 'val' } },
      }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'card_action',
      expect.objectContaining({ actionData: { action: 'confirm', data: { key: 'val' } } })
    );
  });

  it('emits "invoke" for generic invoke activities', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({ type: 'invoke', name: 'genericEvent', value: { foo: 'bar' } }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'invoke',
      expect.objectContaining({ name: 'genericEvent' })
    );
  });
});

// ---------------------------------------------------------------------------
// Channel event handlers
// ---------------------------------------------------------------------------

describe('Channel event handlers', () => {
  let chat: ChatInstance;
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    chat = makeChatInstance();
    adapter = makeAdapter(chat);
    await adapter.initialize(chat);
  });

  afterEach(() => vi.restoreAllMocks());

  it('emits "member_added"', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({
        type: 'conversationUpdate',
        membersAdded: [{ id: 'new-user', name: 'New User' }],
        channelData: {},
      }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'member_added',
      expect.objectContaining({ membersAdded: [{ id: 'new-user', name: 'New User' }] })
    );
  });

  it('emits "member_removed"', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({
        type: 'conversationUpdate',
        membersRemoved: [{ id: 'left-user', name: 'Left User' }],
        channelData: {},
      }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'member_removed',
      expect.objectContaining({ membersRemoved: [{ id: 'left-user', name: 'Left User' }] })
    );
  });

  it('emits "channel_created"', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({
        type: 'conversationUpdate',
        channelData: {
          eventType: 'channelCreated',
          channel: { id: 'new-chan', name: 'new-channel' },
          team: { id: 'team-1' },
        },
      }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'channel_created',
      expect.objectContaining({ channel: { id: 'new-chan', name: 'new-channel' } })
    );
  });
});

// ---------------------------------------------------------------------------
// Lifecycle handlers
// ---------------------------------------------------------------------------

describe('Lifecycle handlers', () => {
  let chat: ChatInstance;
  let adapter: TeamsAdapterImpl;

  beforeEach(async () => {
    chat = makeChatInstance();
    adapter = makeAdapter(chat);
    await adapter.initialize(chat);
  });

  afterEach(() => vi.restoreAllMocks());

  it('emits "app_installed"', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({
        type: 'installationUpdate',
        name: 'add',
        channelData: { action: 'add', team: { id: 'team-abc' } },
      }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'app_installed',
      expect.objectContaining({ installedBy: 'user-1' })
    );
  });

  it('emits "app_uninstalled"', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({
        type: 'installationUpdate',
        name: 'remove',
        channelData: { action: 'remove' },
      }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'app_uninstalled',
      expect.objectContaining({ removedBy: 'user-1' })
    );
  });

  it('emits "bot_activity" for every activity', async () => {
    const validateMod = await import('../src/utils/auth-validator.js');
    vi.spyOn(validateMod, 'validateToken').mockResolvedValue(true);

    await adapter.handleWebhook({
      headers: { authorization: 'Bearer token' },
      body: makeActivity({ type: 'message' }),
    });

    expect(chat.emit).toHaveBeenCalledWith(
      'bot_activity',
      expect.objectContaining({ type: 'message' })
    );
  });
});
