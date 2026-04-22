# Outlook Add-in + Vercel AI SDK: Client-Side Tool Execution

> **Summary of research and architecture decisions** for enabling an AI agent (Vercel AI SDK / Next.js) to trigger Office.js APIs inside an Outlook Add-in task pane via the client-side tools pattern.

---

## Problem Statement

We want an AI agent to interact with the Outlook mailbox (read email, flag, reply, show notifications) from a task pane Add-in. The challenge: Office.js APIs only exist in the **browser context** of the Add-in — they cannot be called from the agent's server-side Node.js environment.

The two apps involved:
- **Agent App** — Next.js server (Vercel), hosts the LLM + tool definitions
- **Outlook Add-in** — Browser task pane, hosts Office.js and `useChat`

---

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                   Outlook Add-in (Task Pane)                 │
│  ┌────────────────────────────────────────────────────────┐  │
│  │  useChat({ transport → Agent App /api/chat })          │  │
│  │                                                        │  │
│  │  onToolCall → Office.js executes → addToolOutput()     │  │
│  └────────────────────────────────────────────────────────┘  │
│              ↕ HTTP streaming (text/event-stream)            │
│  ┌────────────────────────────────────────────────────────┐  │
│              Agent App (separate origin, Next.js)            │
│  │  POST /api/chat                                        │  │
│  │  streamText({ tools: { readEmail, flagEmail, ... }})   │  │
│  │  ← tools have NO execute → forwarded to client        │  │
│  └────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────┘
```

---

## How the SDK Ensures Client-Side Execution (Not Server-Side)

This is enforced across **4 layers** in the Vercel AI SDK source code.

### Layer 1 — `isExecutableTool` check in `execute-tool-call.ts`

```typescript
// packages/ai/src/generate-text/execute-tool-call.ts
export async function executeToolCall(...): Promise<ToolOutput<TOOLS> | undefined> {
  const tool = tools?.[toolName];

  if (!isExecutableTool(tool)) {
    return undefined; // ← silently skipped, nothing runs server-side
  }
  // execute() only runs if the tool has an execute function
}
```

**Rule:** A tool defined **without** an `execute` function will never run server-side. `isExecutableTool` returns `false` and the function returns `undefined`. The tool call is streamed to the client as-is.

### Layer 2 — `stream-text.ts` pauses the agentic loop

After each LLM step, `streamText` checks if it can continue. It will **only call the LLM again** once all client tool calls have received outputs:

```typescript
// packages/ai/src/generate-text/stream-text.ts
const clientToolCalls = stepToolCalls.filter(
  toolCall => toolCall.providerExecuted !== true
);
const clientToolOutputs = stepToolOutputs.filter(
  toolOutput => toolOutput.providerExecuted !== true
);

// Continue loop ONLY when all client tools have been resolved:
if (
  clientToolCalls.length > 0 &&
  clientToolCalls.length === clientToolOutputs.length + deniedResponses.length
) {
  // → trigger next LLM step
}
// Otherwise: stream stays open, loop is paused
```

### Layer 3 — `useChat` auto-resubmits when all outputs are filled

`sendAutomaticallyWhen: lastAssistantMessageIsCompleteWithToolCalls` watches the message state in the Add-in. The moment every pending tool call has a corresponding `addToolOutput()`, it automatically POSTs back to the agent route.

```
Add-in browser                         Agent server
──────────────                         ────────────
onToolCall fires          
  Office.js executes      
  addToolOutput() called  
  [all tools resolved?]   
  YES → auto POST         →            streamText loop resumes
                                       clientToolOutputs.length === clientToolCalls.length
                                       → next LLM call
```

### Layer 4 — WorkflowAgent explicitly stops on tools without `execute`

```typescript
// packages/workflow/src/workflow-agent.ts
const pausedToolCalls = nonProviderToolCalls.filter((tc, i) => {
  const tool = (effectiveTools as ToolSet)[tc.toolName];
  return tool && typeof tool.execute !== 'function'; // ← your Office.js tools
});

// If paused tool calls exist → stop loop, return to client
```

---

## The Golden Rule

| Do | Don't |
|---|---|
| ✅ Define tools **without** `execute` on the agent | ❌ Add `execute` to any Office.js tool — it will run in Node.js and crash |
| ✅ Call `addToolOutput()` in `onToolCall` **without** `await` | ❌ `await addToolOutput()` — causes deadlock inside `onToolCall` |
| ✅ Use `sendAutomaticallyWhen: lastAssistantMessageIsCompleteWithToolCalls` | ❌ Forget this — stream stays paused forever |
| ✅ Guard `onToolCall` with `if (toolCall.dynamic) return` first | ❌ Skip this guard — TypeScript type errors on `toolCall.toolName` |

---

## Verified API Correctness

These were validated against Vercel AI SDK source and Office.js documentation:

### Vercel AI SDK (verified against source)

| API | Correct Usage |
|---|---|
| Route response method | `result.toUIMessageStreamResponse()` ✅ (NOT `toDataStreamResponse()`) |
| Messages parsing | `convertToModelMessages(messages)` with body typed as `UIMessage[]` ✅ |
| Tool schema key | `inputSchema: z.object({...})` ✅ (NOT `parameters:`) |
| `onToolCall` guard | `if (toolCall.dynamic) return;` required first ✅ |
| `addToolOutput` | No `await`, called inside `onToolCall` ✅ |

### Office.js (verified against OfficeDev/Office-Add-in-samples)

| API | Status | Notes |
|---|---|---|
| `item.body.getAsync()` | ✅ Valid | Works in read mode |
| `item.subject`, `item.from`, `item.to` | ✅ Valid | Synchronous read-only properties |
| `item.notificationMessages.addAsync()` | ✅ Valid | Works in read mode |
| `item.displayReplyAllForm()` | ✅ Valid | Opens compose window only — does NOT send |
| `item.flagStatus.setAsync()` | ❌ Does NOT exist | Must use Microsoft Graph `PATCH /me/messages/{id}` |
| `item.isRead.setAsync()` | ❌ Does NOT exist | Must use Microsoft Graph `PATCH /me/messages/{id}` |

### Microsoft Graph (for write operations)

Flag/read operations require:
1. REST-compatible item ID via `Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0)`
2. Access token via MSAL NAA (Nested App Auth) — reference: [Outlook-Add-in-SSO-NAA sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA)
3. `PATCH https://graph.microsoft.com/v1.0/me/messages/{restId}` with `{ flag: { flagStatus: "flagged" } }` or `{ isRead: true }`

---

## Agent Route (Next.js)

```typescript
// agent-app/app/api/chat/route.ts
import { openai } from '@ai-sdk/openai';
import { convertToModelMessages, streamText, UIMessage, tool } from 'ai';
import { z } from 'zod';

export async function OPTIONS() {
  return new Response(null, {
    headers: {
      'Access-Control-Allow-Origin': process.env.ADDIN_ORIGIN!,
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    },
  });
}

export async function POST(req: Request) {
  const { messages }: { messages: UIMessage[] } = await req.json();

  const result = streamText({
    model: openai('gpt-4o'),
    system: `You are an Outlook assistant. You can read emails, flag them,
             move them, reply, and navigate the mailbox UI.
             Always confirm before making destructive actions.`,
    messages: await convertToModelMessages(messages),
    tools: {
      // ── No execute = client-side (Office.js) ──────────────────
      getSelectedEmailBody: tool({
        description: 'Get the body text of the currently selected email.',
        inputSchema: z.object({}),
      }),
      getEmailMetadata: tool({
        description: 'Get subject, sender, recipients, and date of selected email.',
        inputSchema: z.object({}),
      }),
      flagEmail: tool({
        description: 'Flag or unflag via Microsoft Graph API.',
        inputSchema: z.object({
          flagged: z.boolean(),
        }),
      }),
      markAsRead: tool({
        description: 'Mark as read/unread via Microsoft Graph API.',
        inputSchema: z.object({ read: z.boolean() }),
      }),
      openReplyForm: tool({
        description: 'Open reply compose window. User must send manually.',
        inputSchema: z.object({ body: z.string() }),
      }),
      showNotification: tool({
        description: 'Show notification in the task pane.',
        inputSchema: z.object({
          message: z.string(),
          type: z.enum(['informationalMessage', 'errorMessage']),
        }),
      }),
    },
  });

  return result.toUIMessageStreamResponse({
    headers: { 'Access-Control-Allow-Origin': process.env.ADDIN_ORIGIN! },
  });
}
```

---

## Add-in Hook (`useOutlookAgent`)

```typescript
// outlook-addin/src/taskpane/useOfficeTools.ts
import { useChat } from '@ai-sdk/react';
import { DefaultChatTransport, lastAssistantMessageIsCompleteWithToolCalls } from 'ai';

export function useOutlookAgent() {
  const chat = useChat({
    transport: new DefaultChatTransport({
      api: `${process.env.REACT_APP_AGENT_URL}/api/chat`,
      credentials: 'include',
    }),
    sendAutomaticallyWhen: lastAssistantMessageIsCompleteWithToolCalls,

    onToolCall: async ({ toolCall }) => {
      // REQUIRED: guard dynamic tools first
      if (toolCall.dynamic) return;

      const handler = officeToolHandlers[toolCall.toolName];
      if (!handler) {
        // No await — avoids deadlock
        chat.addToolOutput({
          tool: toolCall.toolName,
          toolCallId: toolCall.toolCallId,
          state: 'output-error',
          errorText: `No handler for \