# Teams Bot SDK Analysis: `microsoft/teams.ts` vs `vercel/chat` Adapter

> **Date:** 2026-04-06 (revised; original analysis: 2026-03-31)
> **Context:** Evaluating the best SDK approach for a Teams bot that connects to a remote Vercel AI SDK chatbot.
> **Key event:** `vercel/chat` [PR #302](https://github.com/vercel/chat/pull/302) merged 2026-03-31, migrating the Teams adapter from the deprecated `botbuilder` to `@microsoft/teams.apps` — the same core SDK as `microsoft/teams.ts`.

---

## ⚡ What Changed Since 2026-03-31

Key developments that affect this document's recommendations and code samples:

| Change | Impact |
|---|---|
| **`@microsoft/teams.apps` v2.0.6 released** (2026-03-25) | Version row corrected from `0.0.x` → `2.0.x`. Packages are shipping stable 2.x releases; version concern is lower. |
| **`HttpPlugin` formally deprecated** (PR #433, v2.0.6) | `BridgeHttpAdapter` correctly implements the new `IHttpServerAdapter` interface — it was already using the right pattern. No adapter changes needed. |
| **`IsTargeted` moved from `Activity` to `Account`** (PR #477, v2.0.6) | `ActivitySender.send()` now calls `createTargeted()`/`updateTargeted()` when `recipient.isTargeted === true`. Feature 6 ephemeral implementation is updated accordingly. |
| **`app.name` deprecated** (v2.0.6) | Blueprint code correctly uses `this.userName` (not `app.name`). |
| **`HttpStream` confirmed not exported from `@microsoft/teams.apps`** | `vercel/chat` adapter has a TODO confirming this. Feature 1 blueprint updated to use `activitySender.createStream()` (returns `IStreamer`) instead of importing `HttpStream` directly. |
| **Invoke route alias map confirmed** (routes/invoke/index.ts) | `task/fetch` → **`dialog.open`**, `task/submit` → **`dialog.submit`**, `composeExtension/query` → **`message.ext.query`**. Features 3 and 4 blueprints corrected. |
| **Adaptive Cards default version bumped to 1.5** (PR #476) | Impacts `buildSelectCard()` — default version corrected to `"1.5"`. |

---

## Table of Contents

1. [Repository Overview](#1-repository-overview)
2. [Architecture: What They Actually Are](#2-architecture-what-they-actually-are)
3. [Feature Comparison Matrix](#3-feature-comparison-matrix)
4. [Pros & Cons Deep Dive](#4-pros--cons-deep-dive)
5. [Decision: Recommended Approach](#5-decision-recommended-approach)
6. [Hybrid Approach: Missing Feature Implementations](#6-hybrid-approach-missing-feature-implementations)
   - [Feature 1: Native Teams HTTP Streaming](#feature-1-native-teams-http-streaming)
   - [Feature 2: Teams SSO / OAuth](#feature-2-teams-sso--oauth)
   - [Feature 3: Task Modules (Modals)](#feature-3-task-modules-modals)
   - [Feature 4: Slash Commands (Messaging Extensions)](#feature-4-slash-commands-messaging-extensions)
   - [Feature 5: addReaction / removeReaction](#feature-5-addreaction--removereaction)
   - [Feature 6: Ephemeral Messages](#feature-6-ephemeral-messages)
   - [Feature 7: Select Menus](#feature-7-select-menus)
7. [Complete Hybrid Adapter Blueprint](#7-complete-hybrid-adapter-blueprint)
8. [Critical Timing: hook() Ordering](#8-critical-timing-hook-ordering)
9. [Risk Register](#9-risk-register)
10. [Recommended Factory Pattern](#10-recommended-factory-pattern)

---

## 1. Repository Overview

| | `microsoft/teams.ts` | `vercel/chat` |
|---|---|---|
| **Repo** | [github.com/microsoft/teams.ts](https://github.com/microsoft/teams.ts) | [github.com/vercel/chat](https://github.com/vercel/chat) |
| **Description** | Official Microsoft Teams TypeScript SDK suite | Unified TypeScript SDK for chat bots across platforms (Slack, Teams, Discord, etc.) |
| **Language** | 99.9% TypeScript | 92.2% TypeScript |
| **Stars** | ~81 | ~1596 |
| **Key packages** | `@microsoft/teams.apps`, `.api`, `.ai`, `.cards`, `.graph`, `.openai`, `.dev` | `chat` (core), `@chat-adapter/teams`, `@chat-adapter/slack`, etc. |
| **Teams adapter dependency** | _is_ the SDK | Depends on `@microsoft/teams.apps` v2 (since PR #302) |
| **Primary target** | Long-running Express/Node server | Serverless / Next.js / Vercel edge functions |
| **Package version** | **`2.0.6`** (stable 2.x series; `0.0.0` in source is an nbgv placeholder) | `4.23.0` (`@chat-adapter/teams`) |

---

## 2. Architecture: What They Actually Are

### `microsoft/teams.ts` — 16-Package Monorepo

```
@microsoft/teams.apps       ← Core App class, event routing, HttpServer, middleware, OAuth
@microsoft/teams.api        ← Activity types, REST API client, Bot Framework types
@microsoft/teams.ai         ← Memory, prompts, citations, function definitions, models
@microsoft/teams.cards      ← Type-safe Adaptive Card builder (default version now 1.5)
@microsoft/teams.graph      ← Graph API typed client
@microsoft/teams.graph-endpoints ← Graph endpoint definitions (typed)
@microsoft/teams.openai     ← Azure OpenAI / OpenAI direct integration
@microsoft/teams.botbuilder ← BotBuilder compatibility shim
@microsoft/teams.dev        ← Dev tunnel, local testing utilities
@microsoft/teams.devtools   ← Debug panel
@microsoft/teams.client     ← Teams JS client SDK
@microsoft/teams.common     ← Shared utilities (logging, storage, http, events)
```

**v2.0.6 HTTP architecture note:** `HttpPlugin` is now **deprecated**. The new pattern is `HttpServer` (always present as `app.server`) + a pluggable `IHttpServerAdapter`. The `ExpressAdapter` is the default; custom adapters implement a single `registerRoute()` method. `app.name` is also deprecated — use your own label instead.

### `vercel/chat` Teams Adapter — After PR #302

The adapter is now a **thin serverless bridge** over `@microsoft/teams.apps`:

```
vercel/chat TeamsAdapter
├── this.app              ← @microsoft/teams.apps App instance (the real teams.ts core)
├── this.bridgeAdapter    ← BridgeHttpAdapter (implements the official IHttpServerAdapter)
├── this.graphReader      ← TeamsGraphReader (Graph API message history)
└── implements Adapter<TeamsThreadId, unknown>  ← vercel/chat cross-platform interface
```

**The BridgeHttpAdapter pattern** — how serverless works:

```typescript
// BridgeHttpAdapter implements IHttpServerAdapter (the official v2.0.6 interface,
// replacing the now-deprecated HttpPlugin). It captures the route handler that
// App.initialize() registers and exposes dispatch() for Next.js API routes.
class BridgeHttpAdapter implements IHttpServerAdapter {
  registerRoute(_method, _path, handler) {
    this.handler = handler; // captured once during app.initialize()
  }
  async dispatch(request: Request, options?: WebhookOptions): Promise<Response> {
    // Bridges Web API Request → teams.ts internal handler → Web API Response
    const serverResponse = await this.handler({ body: parsedBody, headers });
    return new Response(JSON.stringify(serverResponse.body), { status: serverResponse.status });
  }
}
```

**Key insight from PR #302** (authored by `@heyitsaamir`, a `teams.ts` core maintainer):
> "We noticed that y'all are using BotFramework in the adapter here which is actually now deprecated. We took the initiative to migrate you guys over to TeamsSDK."

Wins listed in the PR:
- Better type safety for Teams entities and Graph types
- Less setup code
- Reaction support added

**Note:** `vercel/chat` was already using `IHttpServerAdapter` before v2.0.6 formalized the concept. PR #433 validated their approach by making it the official SDK pattern. The adapter is more architecturally forward-looking than originally expected.

---

## 3. Feature Comparison Matrix

### Core Messaging

| Feature | `teams.ts` direct | `vercel/chat` adapter |
|---|---|---|
| Post message | ✅ | ✅ |
| Edit message | ✅ | ✅ |
| Delete message | ✅ | ✅ |
| File uploads | ✅ | ✅ (data URI encoding) |
| Typing indicator | ✅ | ✅ |
| Native HTTP streaming | ✅ `ctx.stream` / `HttpStream` (internal) | ❌ post+edit fallback (TODO in adapter source) |
| DMs / openDM | ✅ | ✅ |

### Rich Content

| Feature | `teams.ts` direct | `vercel/chat` adapter |
|---|---|---|
| Adaptive Cards | ✅ (type-safe `@microsoft/teams.cards`, default v1.5) | ✅ (JSON converter) |
| Buttons (Action.Submit) | ✅ | ✅ |
| Link buttons (Action.OpenUrl) | ✅ | ✅ |
| Select menus (Input.ChoiceSet) | ✅ | ❌ not implemented |
| Tables | ✅ | ✅ GFM |
| Modals (Dialogs: `dialog.open`/`dialog.submit`) | ✅ | ❌ not implemented |
| Targeted/ephemeral messages | ✅ `recipient.isTargeted` (⚠️ experimental) | ❌ DM fallback only |

### Conversations & Identity

| Feature | `teams.ts` direct | `vercel/chat` adapter |
|---|---|---|
| Mentions detection | ✅ | ✅ |
| Receive reactions | ✅ | ✅ |
| Add reactions | ❌ (Teams platform limit — needs delegated token) | ❌ NotImplementedError |
| OAuth / SSO sign-in | ✅ full pipeline | ❌ not implemented |
| Slash commands (`message.ext.query`) | ✅ | ❌ not implemented |
| Multi-platform support | ❌ Teams only | ✅ Slack, Teams, Discord, GChat, etc. |

### AI & Developer Experience

| Feature | `teams.ts` direct | `vercel/chat` adapter |
|---|---|---|
| Vercel AI SDK integration | ⚠️ manual wiring | ✅ native via `thread.stream()` |
| Serverless/Next.js deployment | ⚠️ custom wiring needed | ✅ native |
| State management | Manual | ✅ Redis/Postgres/in-memory, distributed locks |
| Concurrency strategies | Manual | ✅ queue/debounce/concurrent |
| Streaming (Vercel AI SDK) | Manual pipe | ✅ `thread.stream(textStream)` |
| Graph API history | ✅ | ✅ via `fetchMessages()` |

### API Stability & Maintenance

| | `teams.ts` direct | `vercel/chat` adapter |
|---|---|---|
| Maintainer | Microsoft (official) | Vercel |
| Package version | **`2.0.6`** (stable 2.x; regular monthly releases) | `4.23.0` (stable) |
| Active development | ✅ daily commits | ✅ daily commits |
| BotFramework dependency | ❌ none (migrated away) | ❌ none (since PR #302) |
| HTTP pattern | `HttpServer` + `IHttpServerAdapter` (HttpPlugin deprecated) | `BridgeHttpAdapter` implements `IHttpServerAdapter` ✅ |

---

## 4. Pros & Cons Deep Dive

### Option A: `microsoft/teams.ts` Directly

**Pros:**

1. **First-party, authoritative** — Microsoft maintainers, official Teams platform SDK
2. **Full Teams feature depth** — Native streaming (`ctx.stream`), Adaptive Card type-safe builder (v1.5), OAuth/SSO, tabs, config tabs, meeting lifecycle, proactive messaging, MCP plugin
3. **Rich AI tooling** — `@microsoft/teams.ai` with memory, citations, function definitions; `@microsoft/teams.openai` for Azure OpenAI
4. **No abstraction overhead** — Direct access to `ctx.activity`, full `activity.channelData`, all Teams-specific invoke types
5. **Vercel AI SDK** — Wire `streamText()` directly into `ctx.stream.emit()` without normalization layer
6. **OAuth/SSO built-in** — `app.event('signin', ...)`, `ctx.signin()`, `ctx.signout()` — fully typed
7. **Dev tooling** — `@microsoft/teams.dev` tunnel, `@microsoft/teams.devtools` debug panel
8. **Plugin ecosystem** — MCP plugin, custom `IPlugin` interface, DI container
9. **Stable release cadence** — v2.0.x with monthly releases and a clear changelog

**Cons:**

1. **Teams-only** — No future path to Slack/GChat/Discord without full rewrite
2. **Not serverless-native** — Designed for long-running Express servers; Next.js requires custom `IHttpServerAdapter` (PR #433 made this easier but still requires wiring)
3. **Self-managed state** — Distributed locks, subscription state, concurrency management are your problem
4. **Smaller community** — ~81 stars; fewer examples, less community knowledge

### Option B: `vercel/chat` Teams Adapter

**Pros:**

1. **Serverless-first** — `BridgeHttpAdapter` is purpose-built for Next.js/Vercel edge functions
2. **Vercel AI SDK is the target** — `thread.stream(textStream)` maps directly to `streamText()` output
3. **Multi-platform foundation** — Add Slack/GChat/Discord adapters with the same handler code
4. **State management included** — Subscriptions, distributed locks, Redis/Postgres/in-memory out of the box
5. **Concurrency strategies** — Queue/debounce/concurrent message handling (critical for AI bots under load)
6. **Now backed by real `teams.ts`** — PR #302 means running on `@microsoft/teams.apps` v2 under the hood
7. **Normalized AI handoff** — `Thread.post()`, `Thread.stream()` are clean interfaces for AI SDK responses

**Cons:**

1. **Abstraction layer costs** — Lose direct access to Teams-specific features; reach-through requires `raw` or subclassing
2. **Teams features second-class** — Native streaming, Modals, Slash commands, SSO all not yet implemented (adapter has an explicit TODO for streaming)
3. **Very new** — Created December 2025; Teams adapter stable since PR #302 (2026-03-31)
4. **Thin adapter = thin coverage** — ~1000 lines implementing ~15 Teams things vs 27k+ lines across 16 packages in `teams.ts`
5. **Streaming is post+edit polling** — Not the real Teams streaming protocol; visible flicker on long AI responses

---

## 5. Decision: Recommended Approach

### **Use `vercel/chat` with `@chat-adapter/teams` as the foundation, extended with a `HybridTeamsAdapter` subclass.**

**Why:**

Your primary constraints are:
- **Deployment target:** Vercel/Next.js serverless → `vercel/chat` wins outright
- **AI runtime:** Vercel AI SDK → `thread.stream()` is the native integration point
- **Teams platform:** PR #302 resolved the concern about deprecated infrastructure — you're now on `@microsoft/teams.apps` v2 either way

The "real Teams SDK" concern is now a non-issue. Both options use the same underlying `@microsoft/teams.apps` core. The question is purely about the abstraction layer on top.

**When to choose `teams.ts` directly instead:**
- You need native Teams HTTP streaming _today_ (not post+edit)
- You need Teams SSO/OAuth for user identity
- You need task modules (modals) or messaging extensions (slash commands)
- You are **certain** this is Teams-only forever
- You are deploying on a long-running server (Azure App Service, AKS), not serverless

**The hybrid path:** Since `TeamsAdapter` internally uses `@microsoft/teams.apps`, you can subclass it to access the underlying `App` instance and register Teams-native handlers for features not yet exposed by the adapter. This gives you `vercel/chat`'s serverless/AI ergonomics + `teams.ts`'s full Teams feature set.

---

## 6. Hybrid Approach: Missing Feature Implementations

### The Foundational Escape Hatch

All hybrid patterns rely on subclassing `TeamsAdapter` to expose the internal `App` instance:

```typescript
// hybrid-teams-adapter.ts
import { TeamsAdapter, type TeamsAdapterConfig } from "@chat-adapter/teams";
import { App } from "@microsoft/teams.apps";

export class HybridTeamsAdapter extends TeamsAdapter {
  // All hybrid features flow from this single accessor
  get teamsApp(): App {
    return (this as any).app as App;
  }

  constructor(config: TeamsAdapterConfig = {}) {
    super(config);
    // Register additional teams.ts event handlers here, BEFORE initialize()
  }
}
```

> **Why this works:** `TeamsAdapter` declares `private readonly app: App` in its constructor. The field is `private` (TypeScript-only, not JS-private), so `(this as any).app` accesses it. `App` is an unsealed class and the `IHttpServerAdapter` bridge is already wired in. Subclassing gives clean access without monkeypatching. Pin `@chat-adapter/teams` to an exact version and add a test asserting `(adapter as any).app instanceof App` to catch any field renames.

---

### Feature 1: Native Teams HTTP Streaming

**Problem:** The current `TeamsAdapter.stream()` uses post+edit polling (sends a message, then edits it for every chunk). This causes visible flickering and violates the Teams streaming UX contract. The adapter source has a TODO comment: `// TODO: Use native HttpStream for DMs once @microsoft/teams.apps exports it.`

**The real protocol** (`HttpStream` internal class):
1. `TypingActivity` with `channelData: { streamType: 'streaming', streamSequence: N }` for each chunk — shows a live "typing bubble" that fills in
2. Final `MessageActivity` with `.addStreamFinal()` — replaces the typing bubble in-place, atomic delivery
3. `update("Thinking...")` — informative status messages while AI processes

**Important: `HttpStream` is NOT exported from `@microsoft/teams.apps` public API** (confirmed by reading the `http/index.ts` exports). The `vercel/chat` adapter's TODO comment reflects this. The correct approach uses `ActivitySender.createStream()` which returns the exported `IStreamer` interface:

```typescript
// hybrid-teams-adapter.ts
// IStreamer IS exported from @microsoft/teams.apps (via types/streamer.ts → index.ts)
import type { IStreamer } from "@microsoft/teams.apps";
import type { ConversationReference } from "@microsoft/teams.api";

export class HybridTeamsAdapter extends TeamsAdapter {
  get teamsApp(): App { return (this as any).app as App; }

  override async stream(
    threadId: string,
    textStream: AsyncIterable<string | StreamChunk>,
    _options?: StreamOptions
  ): Promise<RawMessage<unknown>> {
    const { conversationId, serviceUrl } = decodeThreadId(threadId);

    // Build the ConversationReference for this thread
    const ref: ConversationReference = {
      channelId: "msteams",
      serviceUrl,
      bot: { id: this.teamsApp.id ?? "", name: this.userName, role: "bot" },
      conversation: { id: conversationId, conversationType: "personal" },
    };

    // activitySender.createStream() creates an HttpStream internally and
    // returns IStreamer — the exported interface with emit(), update(), close()
    // This is the safe approach that avoids importing the non-exported HttpStream class.
    const activitySender = (this.teamsApp as any).activitySender;
    const httpStream: IStreamer = activitySender.createStream(ref);

    for await (const chunk of textStream) {
      const text = typeof chunk === "string" ? chunk
        : chunk.type === "markdown_text" ? chunk.text : "";
      if (text) httpStream.emit(text); // rate-limited (500ms batching, 10 items/pass)
    }

    const result = await httpStream.close(); // sends final MessageActivity with streamFinal
    return { id: (result as any)?.id ?? "", threadId, raw: result };
  }
}
```

**What you gain:**
- True streaming UX — Teams shows a live typing bubble that fills in character-by-character
- `httpStream.update("Thinking...")` — informative status while the AI processes
- Rate limiting handled internally (500ms batching, 10 items/pass)
- Atomic final delivery — no flicker

**Constraint:** Requires Vercel Edge Runtime + adequate timeout:

```typescript
// app/api/webhooks/teams/route.ts
export const runtime = 'edge';   // Required for streaming
export const maxDuration = 60;   // Vercel Pro: up to 300s
```

---

### Feature 2: Teams SSO / OAuth

**Problem:** `TeamsAdapter` has no OAuth surface. `teams.ts` has a complete pipeline:
- `ctx.signin()` — initiates the OAuth card flow
- `ctx.signout()` — token revocation
- `app.event('signin', ...)` — fires after successful token exchange (typed in `IEvents`)
- `signin/tokenExchange` and `signin/verifyState` invoke handlers (registered automatically in `App` constructor)

**Implementation:**

```typescript
// HybridTeamsAdapterConfig extends TeamsAdapterConfig with the OAuth connection name
// so it can be passed at App construction time (options.oauth is readonly after init).
export interface HybridTeamsAdapterConfig extends TeamsAdapterConfig {
  oauthConnectionName?: string;
}

export class HybridTeamsAdapter extends TeamsAdapter {
  private _tokenStore = new Map<string, string>(); // userId → userToken

  constructor(config: HybridTeamsAdapterConfig = {}) {
    // Pass oauthConnectionName through to the App constructor.
    // TeamsAdapter.toAppOptions() does not handle oauth, so we extend the
    // App options via a protected override approach: create the adapter with
    // the standard config, then the hookSSO() method registers the event listener.
    // The App's oauth config is set by patching the readonly options reference
    // before initialize() is called (TypeScript readonly, not Object.freeze()).
    super(config);
    if (config.oauthConnectionName) {
      // options is readonly in TypeScript but not sealed at runtime.
      // Patch it before app.initialize() is called by Chat constructor.
      (this.teamsApp as any).options = {
        ...(this.teamsApp as any).options,
        oauth: { defaultConnectionName: config.oauthConnectionName },
      };
    }
  }

  hookSSO(
    connectionName: string,
    onSignIn: (userId: string, token: string) => Promise<void>
  ): this {
    // app.event('signin') is fully typed via IEvents in @microsoft/teams.apps
    this.teamsApp.event("signin", async (ctx) => {
      const token = (ctx as any).token?.token ?? "";
      this._tokenStore.set((ctx as any).activity.from.id, token);
      await onSignIn((ctx as any).activity.from.id, token);
    });
    return this;
  }

  getUserToken(userId: string): string | undefined {
    return this._tokenStore.get(userId);
  }

  async sendSignInCard(threadId: string, opts?: { text?: string }): Promise<void> {
    const { conversationId } = decodeThreadId(threadId);
    const connectionName = this.teamsApp.oauth?.defaultConnectionName ?? "graph";

    await this.teamsApp.send(conversationId, {
      type: "message",
      attachments: [{
        contentType: "application/vnd.microsoft.card.oauth",
        content: {
          text: opts?.text ?? "Please sign in",
          connectionName,
          buttons: [{ type: "signin", title: "Sign In" }],
        },
      }],
    });
  }
}
```

**Key insight:** `teams.ts` registers `signin/tokenExchange` and `signin/verifyState` route handlers in `App`'s constructor _automatically_. Since `TeamsAdapter` calls `new App(...)`, these are already wired. The `BridgeHttpAdapter` routes all invokes through the same handler chain — SSO invokes work automatically.

**Usage in bot:**

```typescript
bot.onNewMention(async (thread, message) => {
  const token = teamsAdapter.getUserToken(message.author.userId);
  if (!token) {
    await teamsAdapter.sendSignInCard(thread.id, { text: "Sign in to access your data" });
    return;
  }
  // Use token for Graph API calls with delegated user permissions
  const result = streamText({ model: openai("gpt-4o"), messages: [...] });
  await thread.stream(result.textStream);
});
```

---

### Feature 3: Task Modules (Dialogs/Modals)

**Problem:** Teams modals use `task/fetch` (open modal) and `task/submit` (submit modal) invokes. `TeamsAdapter` has no handlers for these.

**Critical: Invoke name aliases.** `teams.ts` maps Teams invoke names to dot-notation aliases via `InvokeAliases` in `routes/invoke/index.ts`. You **must** use the alias, not the raw Teams name:
- `task/fetch` → **`dialog.open`**
- `task/submit` → **`dialog.submit`**

Using the wrong event name (e.g. `"task.fetch"`) will silently never fire.

**Implementation:**

```typescript
export class HybridTeamsAdapter extends TeamsAdapter {
  private _modalHandlers = new Map<string, ModalHandler>();

  // Call BEFORE new Chat({...}) — registers dialog.open + dialog.submit handlers
  hookModals(): this {
    const app = this.teamsApp;

    // Teams sends task/fetch when a button with { msteams: { type: "task/fetch" } } is clicked
    // teams.ts alias: "dialog.open" (not "task.fetch"!)
    app.on("dialog.open", async (ctx) => {
      const data = (ctx.activity.value as any)?.data ?? {};
      const handler = this._modalHandlers.get(data.modalId);
      if (!handler) return { status: 404 };

      const threadId = encodeThreadId({
        conversationId: ctx.activity.conversation.id,
        serviceUrl: ctx.activity.serviceUrl,
      });
      const result = await handler(data, ctx.activity.from.id, threadId);
      return {
        status: 200,
        body: {
          task: {
            type: "continue",
            value: {
              title: result.title,
              height: result.height ?? "medium",
              width: result.width ?? "medium",
              // Either card (Adaptive Card) or url (iframe)
              ...(result.url ? { url: result.url } : {
                card: {
                  contentType: "application/vnd.microsoft.card.adaptive",
                  content: result.card,
                },
              }),
            },
          },
        },
      };
    });

    // Teams sends task/submit when the modal's Action.Submit is clicked
    // teams.ts alias: "dialog.submit" (not "task.submit"!)
    app.on("dialog.submit", async (ctx) => {
      const data = (ctx.activity.value as any)?.data ?? {};
      const handler = this._modalHandlers.get(data.modalId);
      if (!handler) return { status: 200, body: { task: null } }; // null closes the modal

      const threadId = encodeThreadId({
        conversationId: ctx.activity.conversation.id,
        serviceUrl: ctx.activity.serviceUrl,
      });
      const result = await handler(data, ctx.activity.from.id, threadId);
      return {
        status: 200,
        body: { task: result.nextCard ? { type: "continue", value: result.nextCard } : null },
      };
    });

    return this;
  }

  registerModal(modalId: string, handler: ModalHandler): this {
    this._modalHandlers.set(modalId, handler);
    return this;
  }

  // Helper: generate the Adaptive Card action that triggers a modal
  createOpenModalAction(modalId: string, buttonTitle: string, data?: Record<string, unknown>) {
    return {
      type: "Action.Submit",
      title: buttonTitle,
      data: {
        msteams: { type: "task/fetch" }, // This tells Teams to open a modal
        modalId,
        ...data,
      },
    };
  }
}
```

The `Action.Submit` fires the existing `handleAdaptiveCardAction` → `chat.processAction()` pipeline in `TeamsAdapter` — no adapter changes needed. The selected values come through `actionData[inputId]` in the `onAction` handler:

```typescript
bot.onAction("select_submit", async (action) => {
  const selectedValue = (action.raw as any).value?.action?.data?.mySelectId;
  // process selection...
});
```

---

### Feature 4: Slash Commands (Messaging Extensions)

**Problem:** Slash commands in Teams are Messaging Extensions — `composeExtension/query` invokes. Not handled in `TeamsAdapter`.

**Critical: Invoke name alias.** `composeExtension/query` maps to **`message.ext.query`** in `teams.ts` (not `"composeExtension/query"`). Using the raw Teams name will silently never fire.

**Implementation:**

```typescript
export class HybridTeamsAdapter extends TeamsAdapter {
  private _slashCommandHandlers = new Map<string, SlashCommandHandler>();

  hookSlashCommands(): this {
    // teams.ts alias: "message.ext.query" (not "composeExtension/query"!)
    this.teamsApp.on("message.ext.query", async (ctx) => {
      const commandId = (ctx.activity.value as any)?.commandId as string;
      const params: Record<string, string> = {};
      for (const p of ((ctx.activity.value as any)?.parameters ?? [])) params[p.name] = p.value;

      const handler = this._slashCommandHandlers.get(commandId);
      if (!handler) {
        return { status: 200, body: { composeExtension: { type: "message", text: `Unknown command: ${commandId}` } } };
      }

      const result = await handler(params.query ?? "", params, ctx.activity.from.id);
      return { status: 200, body: { composeExtension: result } };
    });

    return this;
  }

  registerSlashCommand(commandId: string, handler: SlashCommandHandler): this {
    this._slashCommandHandlers.set(commandId, handler);
    return this;
  }
}
```

**App manifest entry** (required alongside the code):

```json
{
  "composeExtensions": [{
    "botId": "YOUR_BOT_ID",
    "commands": [{
      "id": "ask",
      "type": "query",
      "title": "Ask the AI",
      "description": "Query the AI knowledge base",
      "parameters": [{
        "name": "query",
        "title": "Query",
        "description": "What to ask"
      }]
    }]
  }]
}
```

---

### Feature 5: addReaction / removeReaction

**Problem:** Both methods throw `NotImplementedError`. The Teams Bot API does not support bot-initiated reactions. The Teams Graph API supports reactions but only via a **delegated user token** (requires SSO).

**Platform reality:** Bot reactions are a Teams platform limitation, not an SDK gap. `addReaction` can only work with a user's delegated Graph token.

**Implementation (requires Feature 2 — SSO):**

```typescript
const TEAMS_EMOJI_MAP: Record<string, string> = {
  thumbsup: "like", heart: "heart", laugh: "laugh",
  surprised: "wow", sad: "sad", angry: "angry",
};

export class HybridTeamsAdapter extends TeamsAdapter {
  override async addReaction(
    threadId: string,
    messageId: string,
    emoji: EmojiValue | string
  ): Promise<void> {
    const emojiName = typeof emoji === "string" ? emoji : emoji.name;
    const teamsEmoji = TEAMS_EMOJI_MAP[emojiName] ?? emojiName;
    const { conversationId } = decodeThreadId(threadId);

    const userToken = [...this._tokenStore.values()][0];
    if (!userToken) {
      throw new Error("addReaction requires a delegated user token. Call hookSSO() and ensure a user has signed in.");
    }

    // DM/group chat path
    await fetch(
      `https://graph.microsoft.com/v1.0/chats/${conversationId}/messages/${messageId}/setReaction`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${userToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({ reactionType: teamsEmoji }),
      }
    );
  }

  override async removeReaction(
    threadId: string,
    messageId: string,
    emoji: EmojiValue | string
  ): Promise<void> {
    const emojiName = typeof emoji === "string" ? emoji : emoji.name;
    const teamsEmoji = TEAMS_EMOJI_MAP[emojiName] ?? emojiName;
    const { conversationId } = decodeThreadId(threadId);

    const userToken = [...this._tokenStore.values()][0];
    if (!userToken) throw new Error("removeReaction requires a delegated user token.");

    await fetch(
      `https://graph.microsoft.com/v1.0/chats/${conversationId}/messages/${messageId}/unsetReaction`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${userToken}`, "Content-Type": "application/json" },
        body: JSON.stringify({ reactionType: teamsEmoji }),
      }
    );
  }
}
```

---

### Feature 6: Ephemeral / Targeted Messages

**Problem:** Teams supports messages visible only to specific users in a group chat via the `isTargeted: true` recipient flag (`Account.isTargeted`). The adapter currently falls back to DMs.

**What changed in v2.0.6 (PR #477 + ActivitySender):** `isTargeted` moved from Activity to Account (the `recipient` object). `ActivitySender.send()` now checks `activity.recipient?.isTargeted === true` and automatically calls the targeted API endpoint (`createTargeted()`/`updateTargeted()`). This means `App.send()` also handles it correctly.

**Note:** Targeted messaging is marked as ⚠️ **experimental/preview** in v2.0.6. The Teams platform requires the Adaptive Cards version 1.5+ and the `supportsChannelFeatures` manifest field.

**Implementation:**

```typescript
export class HybridTeamsAdapter extends TeamsAdapter {
  async postEphemeral(
    threadId: string,
    text: string,
    recipientUserId: string
  ): Promise<void> {
    const { conversationId } = decodeThreadId(threadId);
    const { MessageActivity } = await import("@microsoft/teams.api");

    const activity = new MessageActivity(text);
    activity.textFormat = "markdown";
    // recipient.isTargeted is the Teams targeted message flag (Account property, v2.0.6+)
    // ActivitySender.send() will route to createTargeted() automatically.
    (activity as any).recipient = { id: recipientUserId, isTargeted: true };

    await this.teamsApp.send(conversationId, activity);
  }
}
```

**Note:** Targeted messages only work in group chats and channels, not 1:1 DMs. The feature is still marked as preview — watch for changes in upcoming `@microsoft/teams.apps` releases.

---

### Feature 7: Select Menus

**Problem:** `TeamsAdapter`'s card converter has no `select` type. In Teams, the equivalent is `Input.ChoiceSet` inside an Adaptive Card.

**Implementation:**

```typescript
// cards-extension.ts
export function buildSelectCard(
  title: string,
  selects: SelectMenuDefinition[],
  submitActionId: string = "select_submit"
): object {
  return {
    type: "AdaptiveCard",
    version: "1.5", // updated: @microsoft/teams.cards now defaults to 1.5 (PR #476)
    body: [
      { type: "TextBlock", text: title, weight: "bolder", wrap: true },
      ...selects.map((select) => ({
        type: "Input.ChoiceSet",
        id: select.id,
        label: select.label,
        placeholder: select.placeholder ?? "Select an option",
        style: select.style ?? "compact",
        isRequired: select.required ?? false,
        choices: select.options.map((opt) => ({ title: opt.label, value: opt.value })),
      })),
    ],
    actions: [{ type: "Action.Submit", title: "Submit", data: { actionId: submitActionId } }],
  };
}

interface SelectMenuDefinition {
  id: string;
  label: string;
  placeholder?: string;
  style?: "compact" | "expanded";
  options: Array<{ label: string; value: string }>;
  required?: boolean;
}
```

The `Action.Submit` fires the existing `handleAdaptiveCardAction` → `chat.processAction()` pipeline in `TeamsAdapter` — no adapter changes needed. The selected values come through `actionData[inputId]` in the `onAction` handler:

```typescript
bot.onAction("select_submit", async (action) => {
  const selectedValue = (action.raw as any).value?.action?.data?.mySelectId;
  // process selection...
});
```

---

## 7. Complete Hybrid Adapter Blueprint

```typescript
// hybrid-teams-adapter.ts
import { TeamsAdapter, type TeamsAdapterConfig, decodeThreadId, encodeThreadId } from "@chat-adapter/teams";
import { App } from "@microsoft/teams.apps";
import type { IStreamer } from "@microsoft/teams.apps";  // IStreamer IS exported (types/streamer.ts)
import type { ConversationReference } from "@microsoft/teams.api";
import type { EmojiValue, RawMessage, StreamChunk, StreamOptions } from "chat";

export interface HybridTeamsAdapterConfig extends TeamsAdapterConfig {
  oauthConnectionName?: string;
}

export class HybridTeamsAdapter extends TeamsAdapter {
  private _tokenStore = new Map<string, string>();
  private _modalHandlers = new Map<string, Function>();
  private _slashCommandHandlers = new Map<string, Function>();

  constructor(config: HybridTeamsAdapterConfig = {}) {
    super(config);
    // Patch oauth config before app.initialize() (called by Chat constructor).
    // options is TypeScript readonly but not Object.freeze()'d, so runtime mutation works.
    if (config.oauthConnectionName) {
      (this.teamsApp as any).options = {
        ...(this.teamsApp as any).options,
        oauth: { defaultConnectionName: config.oauthConnectionName },
      };
    }
  }

  // ── Core escape hatch ────────────────────────────────────────────
  get teamsApp(): App {
    return (this as any).app as App;
  }

  // ── Feature 1: Native Streaming ──────────────────────────────────
  // Note: HttpStream is NOT exported from @microsoft/teams.apps public API.
  // We use activitySender.createStream(ref) which returns IStreamer (IS exported).
  // This avoids importing from an internal path.
  override async stream(
    threadId: string,
    textStream: AsyncIterable<string | StreamChunk>,
    _options?: StreamOptions
  ): Promise<RawMessage<unknown>> {
    const { conversationId, serviceUrl } = decodeThreadId(threadId);
    const ref: ConversationReference = {
      channelId: "msteams",
      serviceUrl,
      bot: { id: this.teamsApp.id ?? "", name: this.userName, role: "bot" },
      conversation: { id: conversationId, conversationType: "personal" },
    };

    // activitySender.createStream() creates HttpStream internally and returns IStreamer
    const activitySender = (this.teamsApp as any).activitySender;
    const httpStream: IStreamer = activitySender.createStream(ref);

    for await (const chunk of textStream) {
      const text = typeof chunk === "string" ? chunk
        : chunk.type === "markdown_text" ? chunk.text : "";
      if (text) httpStream.emit(text);
    }

    const result = await httpStream.close();
    return { id: (result as any)?.id ?? "", threadId, raw: result };
  }

  // ── Feature 2: SSO / OAuth ───────────────────────────────────────
  hookSSO(connectionName: string, onSignIn: (userId: string, token: string) => Promise<void>): this {
    // app.event('signin') is fully typed via IEvents in @microsoft/teams.apps (no 'as any' needed)
    this.teamsApp.event("signin", async (ctx) => {
      const token = (ctx as any).token?.token ?? "";
      this._tokenStore.set((ctx as any).activity.from.id, token);
      await onSignIn((ctx as any).activity.from.id, token);
    });
    return this;
  }

  getUserToken(userId: string): string | undefined { return this._tokenStore.get(userId); }

  async sendSignInCard(threadId: string, opts?: { text?: string }): Promise<void> {
    const { conversationId } = decodeThreadId(threadId);
    await this.teamsApp.send(conversationId, {
      type: "message",
      attachments: [{
        contentType: "application/vnd.microsoft.card.oauth",
        content: {
          text: opts?.text ?? "Please sign in",
          connectionName: this.teamsApp.oauth?.defaultConnectionName ?? "graph",
          buttons: [{ type: "signin", title: "Sign In" }],
        },
      }],
    });
  }

  // ── Feature 3: Modals (task/fetch → dialog.open, task/submit → dialog.submit) ─
  // IMPORTANT: teams.ts maps Teams invoke names to dot-notation aliases.
  // Wrong event name = handler silently never fires.
  hookModals(): this {
    const app = this.teamsApp;

    // "task/fetch" maps to the alias "dialog.open" in teams.ts InvokeAliases
    app.on("dialog.open", async (ctx) => {
      const data = (ctx.activity.value as any)?.data ?? {};
      const handler = this._modalHandlers.get(data.modalId);
      if (!handler) return { status: 404 };
      const threadId = encodeThreadId({
        conversationId: ctx.activity.conversation.id,
        serviceUrl: ctx.activity.serviceUrl,
      });
      const result = await handler(data, ctx.activity.from.id, threadId);
      return { status: 200, body: { task: { type: "continue", value: result } } };
    });

    // "task/submit" maps to the alias "dialog.submit" in teams.ts InvokeAliases
    app.on("dialog.submit", async (ctx) => {
      const data = (ctx.activity.value as any)?.data ?? {};
      const handler = this._modalHandlers.get(data.modalId);
      if (!handler) return { status: 200, body: { task: null } };
      const threadId = encodeThreadId({
        conversationId: ctx.activity.conversation.id,
        serviceUrl: ctx.activity.serviceUrl,
      });
      const result = await handler(data, ctx.activity.from.id, threadId);
      return { status: 200, body: { task: result.nextCard ? { type: "continue", value: result.nextCard } : null } };
    });
    return this;
  }

  registerModal(modalId: string, handler: Function): this {
    this._modalHandlers.set(modalId, handler);
    return this;
  }

  // ── Feature 4: Slash Commands (composeExtension/query → message.ext.query) ─
  // IMPORTANT: "composeExtension/query" maps to "message.ext.query" in teams.ts InvokeAliases.
  hookSlashCommands(): this {
    this.teamsApp.on("message.ext.query", async (ctx) => {
      const commandId = (ctx.activity.value as any)?.commandId as string;
      const params: Record<string, string> = {};
      for (const p of ((ctx.activity.value as any)?.parameters ?? [])) params[p.name] = p.value;
      const handler = this._slashCommandHandlers.get(commandId);
      if (!handler) {
        return { status: 200, body: { composeExtension: { type: "message", text: "Unknown command" } } };
      }
      const result = await handler(params.query ?? "", params, ctx.activity.from.id);
      return { status: 200, body: { composeExtension: result } };
    });
    return this;
  }

  registerSlashCommand(commandId: string, handler: SlashCommandHandler): this {
    this._slashCommandHandlers.set(commandId, handler);
    return this;
  }

  // ── Feature 5: Reactions (requires delegated token via SSO) ─────
  override async addReaction(threadId: string, messageId: string, emoji: EmojiValue | string): Promise<void> {
    const teamsEmoji = this._toTeamsEmoji(typeof emoji === "string" ? emoji : emoji.name);
    const { conversationId } = decodeThreadId(threadId);
    const userToken = [...this._tokenStore.values()][0];
    if (!userToken) throw new Error("addReaction requires a delegated user token. Call hookSSO() and ensure a user has signed in.");
    await fetch(`https://graph.microsoft.com/v1.0/chats/${conversationId}/messages/${messageId}/setReaction`, {
      method: "POST",
      headers: { Authorization: `Bearer ${userToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({ reactionType: teamsEmoji }),
    });
  }

  override async removeReaction(threadId: string, messageId: string, emoji: EmojiValue | string): Promise<void> {
    const teamsEmoji = this._toTeamsEmoji(typeof emoji === "string" ? emoji : emoji.name);
    const { conversationId } = decodeThreadId(threadId);
    const userToken = [...this._tokenStore.values()][0];
    if (!userToken) throw new Error("removeReaction requires a delegated user token.");
    await fetch(`https://graph.microsoft.com/v1.0/chats/${conversationId}/messages/${messageId}/unsetReaction`, {
      method: "POST",
      headers: { Authorization: `Bearer ${userToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({ reactionType: teamsEmoji }),
    });
  }

  // ── Internal helpers ─────────────────────────────────────────────
  private _toTeamsEmoji(name: string): string {
    return ({ thumbsup: "like", heart: "heart", laugh: "laugh", surprised: "wow", sad: "sad", angry: "angry" } as Record<string, string>)[name] ?? name;
  }
}
```

---

## 8. Critical Timing: hook() Ordering

**This is the most important implementation constraint.** All `hookX()` calls **must happen before `new Chat({adapters: ...})`** is constructed, because `Chat` calls `adapter.initialize(chatInstance)` which calls `app.initialize()` which locks the router.

```
Required construction order:
1. new HybridTeamsAdapter(config)    ← App created, router empty
2.   .hookSSO(...)                   ← registers app.event('signin')
3.   .hookModals()                   ← registers app.on('dialog.open') + app.on('dialog.submit')
4.   .hookSlashCommands()            ← registers app.on('message.ext.query')
5.   .registerModal(...)             ← populates modal handler map
6.   .registerSlashCommand(...)      ← populates slash command handler map
7. new Chat({ adapters: { teams: adapter } })  ← triggers adapter.initialize()
                                               → app.initialize() called
                                               → router is now locked
```

If you register hooks after step 7, they silently never fire.

---

## 9. Risk Register

| Risk | Severity | Mitigation |
|---|---|---|
| `(this as any).app` breaks if `TeamsAdapter` renames the private field | Medium | Pin `@chat-adapter/teams` to exact version; add unit test `assert((adapter as any).app instanceof App)` |
| `HttpStream` not in `@microsoft/teams.apps` public exports | Medium | Use `activitySender.createStream(ref)` (returns `IStreamer` which IS exported). Fall back to deep import `@microsoft/teams.apps/dist/http/http-stream.js` if needed. Monitor upstream for official export. |
| `activitySender` field rename on `App` | Low | Field is `protected activitySender: ActivitySender` — stable, but add assertion test |
| Streaming requires Edge Runtime; Vercel function timeout | High | Set `export const runtime = 'edge'` and `export const maxDuration = 60` (Pro) / `300` (Enterprise) |
| **Invoke event name aliases** — using raw Teams names silently never fires | **Critical** | Use `teams.ts` aliases: `dialog.open` (not `task.fetch`), `dialog.submit` (not `task.submit`), `message.ext.query` (not `composeExtension/query`). Verify against `routes/invoke/index.ts` `InvokeAliases` map. |
| Graph API reactions need delegated token, not app token | High | Only expose `addReaction` behind an SSO gate; throw with a clear error if no user token |
| `hookX()` ordering must precede `Chat` construction | High | Enforce with the factory function pattern (see Section 10) |
| `options.oauth` mutation via `as any` | Low | Works at runtime (TypeScript `readonly` doesn't prevent mutation). If this breaks, extend `TeamsAdapterConfig` and modify the adapter's `toAppOptions()` instead. |
| `@microsoft/teams.apps` version drift between adapter's dep and direct dep | Medium | Use `peerDependencies` in your app; ensure both resolve to the same version |
| Targeted messages (Feature 6) are ⚠️ experimental in v2.0.6 | Medium | Test in a staging Teams tenant; feature requires Adaptive Cards v1.5+ and manifest changes |

---

## 10. Recommended Factory Pattern

Wrap the entire setup in a factory function to enforce hook ordering, prevent misuse, and provide a clean API:

```typescript
// create-teams-adapter.ts
import { HybridTeamsAdapter, type HybridTeamsAdapterConfig } from "./hybrid-teams-adapter";

interface HybridAdapterOptions extends HybridTeamsAdapterConfig {
  // SSO — oauthConnectionName now in HybridTeamsAdapterConfig
  onSignIn?: (userId: string, token: string) => Promise<void>;
  // Feature flags
  enableModals?: boolean;
  enableSlashCommands?: boolean;
}

// Factory enforces correct hook registration order before Chat() is constructed
export function createHybridTeamsAdapter(opts: HybridAdapterOptions): HybridTeamsAdapter {
  // Constructor patches oauth if oauthConnectionName provided
  const adapter = new HybridTeamsAdapter(opts);

  // SSO always first (other features may depend on user tokens)
  if (opts.oauthConnectionName) {
    adapter.hookSSO(opts.oauthConnectionName, opts.onSignIn ?? (async () => {}));
  }

  if (opts.enableModals) adapter.hookModals();
  if (opts.enableSlashCommands) adapter.hookSlashCommands();

  // Adapter is now ready. All hooks registered. app.initialize() not yet called.
  return adapter;
}
```

```typescript
// bot.ts — complete example
import { Chat } from "chat";
import { createHybridTeamsAdapter } from "./create-teams-adapter";
import { streamText } from "ai";
import { openai } from "@ai-sdk/openai";

// Step 1: Create adapter with all features enabled (hooks registered here)
const teamsAdapter = createHybridTeamsAdapter({
  appType: "SingleTenant",
  oauthConnectionName: "graph",
  enableModals: true,
  enableSlashCommands: true,
  onSignIn: async (userId, token) => {
    console.log(`User ${userId} authenticated`);
  },
});

// Step 2: Register domain-specific handlers (still before Chat construction)
teamsAdapter
  .registerModal("feedback", async (data, userId) => {
    if (data.submitted) { await processFeedback(userId, data); return { title: "Done" }; }
    return { title: "Feedback", height: "medium", card: { /* Adaptive Card JSON */ } };
  })
  .registerSlashCommand("ask", async (query, _params, userId) => ({
    type: "result",
    attachmentLayout: "list",
    attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: { /* result card */ } }],
  }));

// Step 3: ONLY NOW create the Chat instance (triggers app.initialize())
const bot = new Chat({
  adapters: { teams: teamsAdapter },
});

// Step 4: Register message handlers normally via vercel/chat API
bot.onNewMention(async (thread, message) => {
  const token = teamsAdapter.getUserToken(message.author.userId);

  if (!token) {
    await teamsAdapter.sendSignInCard(thread.id, { text: "Sign in to continue" });
    return;
  }

  // Native streaming via activitySender.createStream() (Feature 1)
  const result = streamText({
    model: openai("gpt-4o"),
    messages: [{ role: "user", content: message.text }],
  });

  await thread.stream(result.textStream); // Uses overridden stream() → IStreamer/HttpStream
});

// app/api/webhooks/teams/route.ts
export const runtime = "edge"; // Required for native streaming
export const maxDuration = 60;
export async function POST(request: Request) {
  return bot.handleWebhook("teams", request);
}
```

---

## References

- [vercel/chat repository](https://github.com/vercel/chat)
- [microsoft/teams.ts repository](https://github.com/microsoft/teams.ts)
- [PR #302: Migrate from BotFramework to TeamsSDK](https://github.com/vercel/chat/pull/302) — merged 2026-03-31
- [teams.ts v2.0.6 release notes](https://github.com/microsoft/teams.ts/releases/tag/v2.0.6) — includes HttpServer, targeted messaging, reactions, configurable endpoint
- [PR #433: Introduce HttpServer + IHttpServerAdapter (deprecates HttpPlugin)](https://github.com/microsoft/teams.ts/pull/433)
- [PR #477: Move IsTargeted from Activity to Account](https://github.com/microsoft/teams.ts/pull/477)
- [PR #483: Make messaging endpoint path configurable](https://github.com/microsoft/teams.ts/pull/483)
- [PR #476: Update Adaptive Cards package (default version 1.5)](https://github.com/microsoft/teams.ts/pull/476)
- [teams.ts HttpStream source (internal, not exported)](https://github.com/microsoft/teams.ts/blob/main/packages/apps/src/http/http-stream.ts)
- [teams.ts ActivitySender + createStream()](https://github.com/microsoft/teams.ts/blob/main/packages/apps/src/activity-sender.ts)
- [teams.ts IStreamer type (exported)](https://github.com/microsoft/teams.ts/blob/main/packages/apps/src/types/streamer.ts)
- [teams.ts invoke route aliases (InvokeAliases)](https://github.com/microsoft/teams.ts/blob/main/packages/apps/src/routes/invoke/index.ts)
- [teams.ts ActivityContext (SSO/signin)](https://github.com/microsoft/teams.ts/blob/main/packages/apps/src/contexts/activity.ts)
- [teams.ts OAuth handlers (onTokenExchange, onVerifyState, onSignInFailure)](https://github.com/microsoft/teams.ts/blob/main/packages/apps/src/app.oauth.ts)
- [teams.ts IHttpServerAdapter interface](https://github.com/microsoft/teams.ts/blob/main/packages/apps/src/http/adapter.ts)
- [vercel/chat TeamsAdapter index.ts](https://github.com/vercel/chat/blob/main/packages/adapter-teams/src/index.ts)
- [vercel/chat BridgeHttpAdapter](https://github.com/vercel/chat/blob/main/packages/adapter-teams/src/bridge-adapter.ts)
