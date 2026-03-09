# @chat/adapter-teamssdk

A Microsoft Teams SDK Adapter for the **vercel/chat** Adapter interface pattern. This package provides a fully-typed, zero-runtime-dependency Teams adapter that handles Bot Framework webhooks, posts messages, manages reactions, opens DMs, routes events, and integrates with the Microsoft Graph API.

---

## Table of Contents

- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
  - [Initialization](#initialization)
  - [Handling Webhooks](#handling-webhooks)
  - [Posting Messages](#posting-messages)
  - [Editing & Deleting Messages](#editing--deleting-messages)
  - [Reactions](#reactions)
  - [Direct Messages](#direct-messages)
  - [Typing Indicators](#typing-indicators)
  - [Fetching Messages & Threads](#fetching-messages--threads)
  - [Thread ID Helpers](#thread-id-helpers)
  - [Teams Event Handling](#teams-event-handling)
  - [Message Formatting](#message-formatting)
  - [Modals](#modals)
- [Error Handling](#error-handling)
- [Security Considerations](#security-considerations)
- [Troubleshooting](#troubleshooting)

---

## Installation

```bash
npm install @chat/adapter-teamssdk
```

> **Requirements**: Node.js ≥ 20. Uses native `fetch` and `Buffer` – no polyfills needed.

---

## Configuration

The adapter accepts a `TeamsAdapterConfig` object. Three authentication methods are supported.

### 1. Client Secret (Password)

```typescript
import { createTeamsAdapter } from '@chat/adapter-teamssdk';

const adapter = createTeamsAdapter({
  appId: process.env.TEAMS_APP_ID!,
  appPassword: process.env.TEAMS_APP_PASSWORD!,
  tenantId: process.env.TEAMS_TENANT_ID,
  enableLogging: true,
});
```

### 2. Certificate-Based Auth

```typescript
const adapter = createTeamsAdapter({
  appId: process.env.TEAMS_APP_ID!,
  appCertificate: {
    thumbprint: process.env.CERT_THUMBPRINT!,    // hex-encoded SHA-1 thumbprint
    privateKey: process.env.CERT_PRIVATE_KEY!,   // PEM-encoded PKCS#8 private key
  },
  tenantId: process.env.TEAMS_TENANT_ID,
});
```

### 3. Federated Identity (Workload Identity)

```typescript
const adapter = createTeamsAdapter({
  appId: process.env.TEAMS_APP_ID!,
  // Federated token injected at runtime (e.g. from Azure Managed Identity)
  // Use getAccessToken({ type: 'federated', appId, federatedToken }) directly
});
```

### Full Config Reference

```typescript
interface TeamsAdapterConfig {
  appId: string;                     // Required – Azure AD App ID
  appPassword?: string;              // Client secret
  appCertificate?: {
    thumbprint: string;
    privateKey: string;
  };
  tenantId?: string;                 // Restricts to single tenant
  allowedTenants?: string[];         // Allowlist of tenant IDs
  enableLogging?: boolean;           // Default: false
  maxRetries?: number;               // Default: 3
  retryDelayMs?: number;             // Base back-off delay (ms). Default: 500
  graphApiBaseUrl?: string;          // Default: https://graph.microsoft.com/v1.0
  botFrameworkApiUrl?: string;       // Default: https://smba.trafficmanager.net/apis
}
```

---

## Usage

### Initialization

```typescript
const chatInstance = {
  emit(event, data) { /* your event bus */ },
  on(event, handler) { /* subscribe */ },
};

await adapter.initialize(chatInstance);
```

### Handling Webhooks

Wire the adapter into your HTTP server. The `handleWebhook` method validates the Bot Framework JWT, parses the activity, and routes events.

```typescript
// Express example
app.post('/api/messages', async (req, res) => {
  const result = await adapter.handleWebhook({
    headers: req.headers as Record<string, string>,
    body: req.body,
    method: req.method,
    url: req.url,
  });
  res.status(result.status).json(result.body ?? {});
});
```

### Posting Messages

```typescript
// Encode a context into a thread ID first
const threadId = adapter.encodeThreadId({
  serviceUrl: 'https://smba.trafficmanager.net/apis',
  tenantId: 'your-tenant-id',
  conversationId: activity.conversation.id,
  teamId: channelData.team?.id,
  channelId: channelData.channel?.id,
});

const sent = await adapter.postMessage(threadId, {
  text: 'Hello from the adapter!',
});
console.log('Sent message ID:', sent.id);
```

### Editing & Deleting Messages

```typescript
await adapter.editMessage(threadId, messageId, {
  text: 'Updated message content',
});

await adapter.deleteMessage(threadId, messageId);
```

### Reactions

Reactions use the Microsoft Graph API. The thread context must include `teamId` and `channelId`.

```typescript
await adapter.addReaction(threadId, messageId, 'like');
await adapter.removeReaction(threadId, messageId, 'like');
```

### Direct Messages

```typescript
// Opens a 1:1 chat with a user and returns an encoded thread ID
const dmThreadId = await adapter.openDM('aad-user-id');

// Send a message to the DM
await adapter.postMessage(dmThreadId, { text: 'Hey, just wanted to check in!' });
```

### Typing Indicators

```typescript
await adapter.startTyping(threadId);
```

### Fetching Messages & Threads

```typescript
// Requires teamId and channelId in the thread context
const messages = await adapter.fetchMessages(threadId);

const singleMessage = await adapter.fetchMessage(threadId, messageId);

// Using composite "teamId:channelId" format
const channelMessages = await adapter.fetchChannelMessages('team-id:channel-id');

const threads = await adapter.listThreads('team-id:channel-id');

const channelInfo = await adapter.fetchChannelInfo('team-id:channel-id');
```

### Thread ID Helpers

Thread IDs are base64url-encoded `TeamsContext` objects.

```typescript
const context = adapter.decodeThreadId(threadId);
// {
//   serviceUrl: 'https://smba.trafficmanager.net/apis',
//   tenantId: 'xxx',
//   conversationId: 'yyy',
//   channelId: 'zzz',
//   teamId: 'aaa',
// }

const channelId = adapter.channelIdFromThreadId(threadId);

const isPersonal = adapter.isDM(threadId); // true for personal/DM conversations
```

### Teams Event Handling

The adapter exposes a `TeamsApp` interface via `adapter.app` for registering event handlers. Handlers receive the raw `TeamsActivity` and the adapter instance.

```typescript
// Respond to any message
adapter.app.$onMessage(async (activity, adapter) => {
  const message = adapter.parseMessage(activity);
  console.log('Received:', message.text);
});

// Respond to @mentions
adapter.app.$onMention(async (activity, adapter) => {
  const threadId = adapter.encodeThreadId({
    serviceUrl: activity.serviceUrl,
    tenantId: activity.conversation.tenantId ?? '',
    conversationId: activity.conversation.id,
  });
  await adapter.postMessage(threadId, { text: 'You mentioned me!' });
});

// DMs
adapter.app.$onDMReceived(async (activity, adapter) => {
  console.log('DM from', activity.from.name);
});

// Reactions
adapter.app.$onReactionAdded(async (activity) => {
  console.log('Reaction added:', activity.reactionsAdded);
});

adapter.app.$onReactionRemoved(async (activity) => {
  console.log('Reaction removed:', activity.reactionsRemoved);
});

// Thread replies
adapter.app.$onThreadReplyAdded(async (activity) => {
  console.log('Reply in thread:', activity.replyToId);
});

// Card actions / Adaptive Card submits
adapter.app.$onCardAction(async (activity) => {
  console.log('Card submitted:', activity.value);
});

// Invoke (task/fetch, compose extensions, etc.)
adapter.app.$onInvoke(async (activity) => {
  console.log('Invoke:', activity.name, activity.value);
});

// Channel / team events
adapter.app.$onChannelCreated(async (activity) => {
  const cd = activity.channelData as any;
  console.log('Channel created:', cd.channel?.name);
});

adapter.app.$onMemberAdded(async (activity) => {
  console.log('Members added:', activity.membersAdded);
});

// App lifecycle
adapter.app.$onAppInstalled(async (activity) => {
  console.log('App installed by', activity.from.name);
});

adapter.app.$onAppUninstalled(async (activity) => {
  console.log('App uninstalled');
});

// Catch-all
adapter.app.$onBotActivity(async (activity) => {
  console.log('Activity received:', activity.type);
});
```

### Message Formatting

```typescript
import type { FormattedContent } from '@chat/adapter-teamssdk';

// Markdown
const md: FormattedContent = {
  type: 'markdown',
  content: '**Bold** and _italic_ with `code` and [link](https://example.com)',
};
console.log(adapter.renderFormatted(md));
// → <strong>Bold</strong> and <em>italic</em> with <code>code</code> and <a href="...">link</a>

// HTML passthrough
const html: FormattedContent = {
  type: 'html',
  content: '<b>Raw HTML</b>',
};

// Adaptive Card (sent as attachment, not text)
import { adaptiveCardToAttachment } from '@chat/adapter-teamssdk';

const cardAttachment = adaptiveCardToAttachment({
  body: [
    { type: 'TextBlock', text: 'Hello from an Adaptive Card!', weight: 'Bolder' },
    { type: 'TextBlock', text: 'Tap the button below.' },
  ],
  actions: [
    { type: 'Action.Submit', title: 'Click me', data: { action: 'clicked' } },
  ],
});

await adapter.postMessage(threadId, {
  text: 'Fallback text for notifications',
  attachments: [{ type: 'card', content: cardAttachment.content, contentType: cardAttachment.contentType }],
});

// Blocks (Slack-style)
const blocks: FormattedContent = {
  type: 'blocks',
  content: '',
  blocks: [
    { type: 'header', text: 'Weekly Report' },
    { type: 'section', text: 'Everything is going well.' },
    { type: 'divider' },
  ],
};
```

### Modals

In Teams, modals (Task Modules) are opened by responding to an `invoke` activity with a `task/fetch` response. The `openModal` method logs the intent; in a real integration you must reply to the invoke activity directly:

```typescript
adapter.app.$onInvoke(async (activity) => {
  if (activity.name === 'task/fetch') {
    // Return a task/fetch response body from your HTTP handler:
    const taskFetchResponse = {
      task: {
        type: 'continue',
        value: {
          title: 'My Modal',
          height: 400,
          width: 500,
          card: adaptiveCardToAttachment({
            body: [{ type: 'Input.Text', id: 'name', label: 'Your name' }],
            actions: [{ type: 'Action.Submit', title: 'Submit' }],
          }),
        },
      },
    };
    // Send taskFetchResponse as the HTTP response body with status 200
  }
});
```

---

## Error Handling

All adapter errors extend `TeamsAdapterError`:

```typescript
import { TeamsAdapterError, TeamsAdapterErrorCode } from '@chat/adapter-teamssdk';

try {
  await adapter.postMessage(threadId, { text: 'Hello' });
} catch (err) {
  if (err instanceof TeamsAdapterError) {
    switch (err.code) {
      case TeamsAdapterErrorCode.UNAUTHORIZED:
        console.error('Auth failed – check appId/appPassword');
        break;
      case TeamsAdapterErrorCode.RATE_LIMITED:
        console.warn('Rate limited – retry later');
        break;
      case TeamsAdapterErrorCode.NOT_FOUND:
        console.error('Conversation not found');
        break;
      default:
        console.error(`Teams error [${err.code}]:`, err.message);
    }
  }
}
```

### Error Codes

| Code | HTTP Status | Description |
|------|-------------|-------------|
| `UNAUTHORIZED` | 401 | Invalid or missing token |
| `FORBIDDEN` | 403 | Insufficient permissions |
| `NOT_FOUND` | 404 | Resource not found |
| `RATE_LIMITED` | 429 | Too many requests |
| `VALIDATION_ERROR` | 400 | Bad request / missing parameters |
| `API_ERROR` | 5xx | Server-side error |
| `UNKNOWN` | 500 | Unexpected error |

The adapter automatically **retries** on transient errors (5xx, rate limits) with exponential back-off. Configure via `maxRetries` and `retryDelayMs`.

---

## Security Considerations

1. **JWT Validation**: Every incoming webhook request is validated against the [Bot Framework JWKS endpoint](https://login.botframework.com/v1/.well-known/openidconfiguration). Both claims (`aud`, `iss`, `exp`) and key ID presence are checked. Full cryptographic RS256 signature verification is performed when the `SubtleCrypto` Web Crypto API is available (Node ≥ 18).

2. **Tenant Isolation**: Set `allowedTenants` to restrict which Azure AD tenants can interact with your bot. This prevents cross-tenant attacks.

3. **Secrets**: Never commit `appPassword` or private keys to source control. Use environment variables or a secrets manager (Azure Key Vault, GitHub Secrets, etc.).

4. **Certificate Auth**: Prefer certificate-based authentication over client secrets in production – certificates are rotatable and do not appear in logs.

5. **Thread ID Encoding**: Thread IDs are **base64url-encoded JSON** (not encrypted). Do not embed sensitive information in the `TeamsContext`.

6. **SSRF via serviceUrl**: The `serviceUrl` from incoming Bot Framework activities is cached and used for outbound calls. Validate that it points to a legitimate Microsoft endpoint in production; do not allow arbitrary values.

7. **Input Validation**: User-supplied text is passed through format converters which escape HTML special characters for `text` type content. Ensure you validate `value` payloads from card actions before acting on them.

---

## Troubleshooting

### `UNAUTHORIZED` on every incoming request

- Ensure your bot is registered in the [Azure Bot Service](https://portal.azure.com) with the correct `appId`.
- Verify the `Authorization: Bearer <token>` header is being forwarded by your HTTP framework (some frameworks strip unknown headers).
- Check that your server clock is accurate – JWT expiry checks are time-sensitive.

### Token acquisition failures

- Check that `appPassword` matches the secret in the Azure Bot registration.
- For certificate auth, ensure the PEM key is PKCS#8 (not PKCS#1). Convert with: `openssl pkcs8 -topk8 -nocrypt -in key.pem -out key-pkcs8.pem`.

### `NOT_FOUND` when posting messages

- The `serviceUrl` in the encoded thread context must match the one received from the Bot Framework.
- Conversation IDs are ephemeral in some scopes – store them from incoming activities rather than constructing them manually.

### Graph API errors

- Graph API calls require a token scoped to `https://graph.microsoft.com/.default`, not the Bot Framework scope. If you encounter permission errors, ensure your Azure AD app registration has the required Graph permissions (`ChannelMessage.Read.All`, `ChannelMessage.Send`, etc.) and admin consent has been granted.

### TypeScript compilation errors

- Ensure `"moduleResolution": "NodeNext"` is set in `tsconfig.json`.
- All imports from this package must use the `.js` extension in TypeScript source (ES module convention).

### Tests failing with `fetch is not defined`

- Tests use `vi.stubGlobal('fetch', ...)` to mock fetch. Ensure you are running vitest ≥ 2.0 with `environment: 'node'`.
