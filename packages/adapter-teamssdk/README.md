# @chat-adapter/teamssdk

Microsoft Teams SDK adapter for [vercel/chat](https://github.com/vercel/chat), built on the
Bot Framework (`botbuilder`) with a **teams.ts-style event-routing abstraction** (`TeamsApp`).

## Architecture

```
handleWebhook(Request)
  → botbuilder CloudAdapter.processActivity
    → TeamsApp.processActivity          ← teams.ts-style event router
      → $onMessage / $onMention / $onCardAction / …
        → chat.processMessage / processAction / processReaction
```

The key difference from `@chat-adapter/teams` is the `TeamsApp` class that sits between
botbuilder and the Chat SDK.  It lets you register typed event handlers with a clean,
composable API (`$onMessage`, `$onMention`, `$onReactionAdded`, etc.) before they are
forwarded to the Chat instance.

## Installation

```bash
npm install @chat-adapter/teamssdk
```

## Quick Start

```ts
import { Chat } from "chat";
import { TeamsSDKAdapter } from "@chat-adapter/teamssdk";

const adapter = new TeamsSDKAdapter({
  appId: process.env.TEAMS_APP_ID,
  appPassword: process.env.TEAMS_APP_PASSWORD,
});

// Optional: attach custom event handlers to the TeamsApp router
adapter.app.$onMention(async (activity) => {
  console.log("Bot was mentioned:", activity.text);
});

adapter.app.$onAppInstalled(async (activity) => {
  console.log("App installed in team:", activity.channelData?.team?.name);
});

const chat = new Chat({ adapter });

// In your webhook handler (e.g. Next.js API route):
export async function POST(request: Request) {
  return adapter.handleWebhook(request);
}
```

## Configuration

```ts
export interface TeamsSDKAdapterConfig {
  /** Microsoft App ID (or set TEAMS_APP_ID env var) */
  appId?: string;
  /** App Password (or set TEAMS_APP_PASSWORD env var) */
  appPassword?: string;
  /** Tenant ID for SingleTenant apps and Graph API (or set TEAMS_APP_TENANT_ID) */
  appTenantId?: string;
  /** "MultiTenant" (default) or "SingleTenant" */
  appType?: "MultiTenant" | "SingleTenant";
  /** Certificate-based auth (alternative to appPassword) */
  certificate?: TeamsAuthCertificate;
  /** Federated workload identity (alternative to appPassword) */
  federated?: TeamsAuthFederated;
  /** Custom logger */
  logger?: Logger;
  /** Bot display name */
  userName?: string;
}
```

### Authentication Methods

**App Password (recommended for development):**
```ts
new TeamsSDKAdapter({ appId: "...", appPassword: "..." })
```

**Certificate:**
```ts
new TeamsSDKAdapter({
  appId: "...",
  appTenantId: "...",
  certificate: {
    certificatePrivateKey: "-----BEGIN RSA PRIVATE KEY-----\n...",
    certificateThumbprint: "abc123...",  // or x5c for SNI
  },
});
```

**Federated Identity (Workload Identity):**
```ts
new TeamsSDKAdapter({
  appId: "...",
  appTenantId: "...",
  federated: { clientId: "managed-identity-client-id" },
});
```

## TeamsApp Event Handlers

The `adapter.app` property exposes the `TeamsApp` event router.  Every handler
receives the raw `botbuilder` `Activity` and `TurnContext`.

| Method | Fires when… |
|---|---|
| `$onMessage(handler)` | Regular channel/group message |
| `$onMention(handler)` | Bot @mentioned |
| `$onThreadReplyAdded(handler)` | Reply posted in a thread |
| `$onDMReceived(handler)` | 1:1 DM message |
| `$onReactionAdded(handler)` | Emoji reaction added |
| `$onReactionRemoved(handler)` | Emoji reaction removed |
| `$onCardAction(handler)` | Adaptive Card button clicked |
| `$onInvoke(handler)` | Task module / general invoke |
| `$onMessageAction(handler)` | Message context-menu action |
| `$onMemberAdded(handler)` | Member joined team/channel |
| `$onMemberRemoved(handler)` | Member left team/channel |
| `$onTeamRenamed(handler)` | Team renamed |
| `$onChannelCreated(handler)` | Channel created |
| `$onChannelRenamed(handler)` | Channel renamed |
| `$onChannelDeleted(handler)` | Channel deleted |
| `$onAppInstalled(handler)` | App installed |
| `$onAppUninstalled(handler)` | App uninstalled |
| `$onBotActivity(handler)` | Any activity (fires first, always) |

All registration methods return `this` for chaining:

```ts
adapter.app
  .$onMessage(handleMessage)
  .$onMention(handleMention)
  .$onCardAction(handleAction);
```

## Adapter Interface

`TeamsSDKAdapter` implements the full `Adapter<TeamsThreadId, unknown>` interface from `chat`:

### Message Operations

```ts
await adapter.postMessage(threadId, "Hello!");
await adapter.editMessage(threadId, messageId, "Updated text");
await adapter.deleteMessage(threadId, messageId);
await adapter.fetchMessages(threadId, { limit: 50 });
await adapter.fetchThread(threadId);
```

### Channel Operations (requires Graph API / appTenantId)

```ts
await adapter.fetchChannelMessages(channelId, { limit: 50 });
await adapter.postChannelMessage(channelId, "Announcement!");
await adapter.fetchChannelInfo(channelId);
await adapter.listThreads(channelId);
```

### Direct Messages

```ts
const dmThreadId = await adapter.openDM(userId);
await adapter.postMessage(dmThreadId, "Hi there!");
```

### Typing Indicator

```ts
await adapter.startTyping(threadId);
```

### Thread ID Utilities

```ts
const threadId = adapter.encodeThreadId({ conversationId, serviceUrl });
const { conversationId, serviceUrl } = adapter.decodeThreadId(threadId);
const channelId = adapter.channelIdFromThreadId(threadId);
const isDM = adapter.isDM(threadId);
```

### Message Formatting

```ts
// Render an AST (FormattedContent) back to Teams markdown
const teamsText = adapter.renderFormatted(message.formatted);

// Parse a raw Teams activity into a normalized Message
const message = adapter.parseMessage(activity);
```

## Rich Messages

### Adaptive Cards

```ts
import { Card, Button, Actions, Fields, Field } from "chat";

await thread.post(
  Card({
    title: "Approval Required",
    subtitle: "Review the request below",
    children: [
      Fields([
        Field({ label: "Requester", value: "Alice" }),
        Field({ label: "Amount", value: "$500" }),
      ]),
      Actions([
        Button({ id: "approve", label: "Approve", style: "primary" }),
        Button({ id: "reject",  label: "Reject",  style: "danger" }),
      ]),
    ],
  })
);
```

### Markdown Formatting

```ts
await thread.post({ markdown: "**Bold**, _italic_, `code`" });
```

## Graph API Features

When `appTenantId` is configured, the adapter automatically sets up a Microsoft Graph client
that enables:

- `fetchMessages()` — load message history from chats and channels
- `fetchChannelMessages()` / `postChannelMessage()` — channel-level operations
- `fetchChannelInfo()` / `listThreads()` — channel metadata

Required Azure AD app permissions (select based on scope):

| Feature | Permission |
|---|---|
| Read chat messages | `ChatMessage.Read.Chat` |
| Read all chats | `Chat.Read.All` |
| Read channels | `ChannelMessage.Read.All` |
| Post to channels | `ChannelMessage.Send` |

## Environment Variables

| Variable | Description |
|---|---|
| `TEAMS_APP_ID` | Microsoft App ID |
| `TEAMS_APP_PASSWORD` | App password |
| `TEAMS_APP_TENANT_ID` | Tenant ID (required for SingleTenant and Graph API) |
