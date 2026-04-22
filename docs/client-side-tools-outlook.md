# Outlook Add-in + Vercel AI SDK: Client-Side Tool Execution

Summary of research and architecture decisions for enabling an AI agent (Vercel AI SDK / Next.js) to trigger Office.js APIs inside an Outlook Add-in task pane via the client-side tools pattern.

---

## Problem Statement

We want an AI agent to interact with the Outlook mailbox (read email, flag, reply, show notifications) from a task pane Add-in. The challenge: Office.js APIs only exist in the browser context of the Add-in and cannot be called from the agent server-side Node.js environment.

The two apps involved:
- Agent App: Next.js server (Vercel), hosts the LLM and tool definitions
- Outlook Add-in: Browser task pane, hosts Office.js and useChat

---

## Architecture

Agent App and Outlook Add-in run as separate web apps.

- The Add-in task pane runs useChat, connecting via HTTP streaming to the Agent App at /api/chat
- onToolCall fires in the browser, executes Office.js, then calls addToolOutput()
- The Agent App runs streamText with tools that have NO execute function, so they are forwarded to the client

---

## How the SDK Ensures Client-Side Execution

This is enforced across 4 layers in the Vercel AI SDK source code.

### Layer 1: isExecutableTool check (execute-tool-call.ts)

Source: packages/ai/src/generate-text/execute-tool-call.ts

```ts
const tool = tools?.[toolName];

if (!isExecutableTool(tool)) {
  return undefined; // silently skipped, nothing runs server-side
}
```

A tool defined without an execute function will never run server-side. isExecutableTool returns false and the function returns undefined. The tool call is streamed to the client as-is.

### Layer 2: stream-text.ts pauses the agentic loop

Source: packages/ai/src/generate-text/stream-text.ts

After each LLM step, streamText checks if it can continue. It will only call the LLM again once all client tool calls have received outputs back from the client.

```ts
const clientToolCalls = stepToolCalls.filter(
  toolCall => toolCall.providerExecuted !== true
);
const clientToolOutputs = stepToolOutputs.filter(
  toolOutput => toolOutput.providerExecuted !== true
);

// Only continue when all client tools resolved:
if (
  clientToolCalls.length > 0 &&
  clientToolCalls.length === clientToolOutputs.length + deniedResponses.length
) {
  // trigger next LLM step
}
// Otherwise: stream stays open, loop paused
```

### Layer 3: useChat auto-resubmits via sendAutomaticallyWhen

sendAutomaticallyWhen: lastAssistantMessageIsCompleteWithToolCalls watches the message state. The moment every pending tool call has a corresponding addToolOutput(), it automatically POSTs back to the agent route to resume the loop.

### Layer 4: WorkflowAgent explicitly stops on tools without execute

Source: packages/workflow/src/workflow-agent.ts

```ts
const pausedToolCalls = nonProviderToolCalls.filter((tc, i) => {
  const tool = (effectiveTools as ToolSet)[tc.toolName];
  return tool && typeof tool.execute !== 'function';
});
// If paused tool calls exist, stop loop and return to client
```

---

## The Golden Rules

DO:
- Define tools without execute on the agent route
- Call addToolOutput() inside onToolCall WITHOUT await (await causes deadlock)
- Use sendAutomaticallyWhen: lastAssistantMessageIsCompleteWithToolCalls
- Guard onToolCall with if (toolCall.dynamic) return; as the very first line

DO NOT:
- Add execute to any Office.js tool - it will run in Node.js where Office.js does not exist
- await addToolOutput() - this deadlocks the onToolCall callback
- Forget sendAutomaticallyWhen - the stream will stay paused forever
- Skip the dynamic guard - causes TypeScript errors on toolCall.toolName

---

## Verified API Correctness

### Vercel AI SDK

- Route response: result.toUIMessageStreamResponse() - NOT toDataStreamResponse()
- Messages: convertToModelMessages(messages) with body typed as UIMessage[]
- Tool schema key: inputSchema: z.object({}) - NOT parameters:
- onToolCall guard: if (toolCall.dynamic) return; must be first
- addToolOutput: no await, called inside onToolCall

### Office.js APIs

Valid in read mode:
- item.body.getAsync() - works
- item.subject, item.from, item.to - synchronous read-only properties
- item.notificationMessages.addAsync() - works
- item.displayReplyAllForm() - opens compose window only, user must send manually

Does NOT exist (common mistakes):
- item.flagStatus.setAsync() - use Microsoft Graph PATCH /me/messages/{id} instead
- item.isRead.setAsync() - use Microsoft Graph PATCH /me/messages/{id} instead

### Microsoft Graph for Write Operations

Steps required:
1. Get REST-compatible item ID: Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0)
2. Get access token via MSAL NAA (Nested App Auth)
3. PATCH https://graph.microsoft.com/v1.0/me/messages/{restId} with body { flag: { flagStatus: "flagged" } } or { isRead: true }

Reference: https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA

---

## Agent Route (Next.js)

```ts
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
    system: 'You are an Outlook assistant. You can read emails, flag them, reply, and show notifications. Always confirm before making destructive actions.',
    messages: await convertToModelMessages(messages),
    tools: {
      // No execute = client-side. Office.js handles these in the browser.
      getSelectedEmailBody: tool({
        description: 'Get the body text of the currently selected email.',
        inputSchema: z.object({}),
      }),
      getEmailMetadata: tool({
        description: 'Get subject, sender, recipients, and date of the selected email.',
        inputSchema: z.object({}),
      }),
      flagEmail: tool({
        description: 'Flag or unflag the selected email via Microsoft Graph.',
        inputSchema: z.object({
          flagged: z.boolean().describe('true to flag, false to clear the flag'),
        }),
      }),
      markAsRead: tool({
        description: 'Mark the selected email as read or unread via Microsoft Graph.',
        inputSchema: z.object({ read: z.boolean() }),
      }),
      openReplyForm: tool({
        description: 'Open the reply-all compose window pre-filled with a body. User must send manually.',
        inputSchema: z.object({ body: z.string() }),
      }),
      showNotification: tool({
        description: 'Show a notification banner in the Outlook task pane.',
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
}
```

---

## Add-in Hook (useOutlookAgent)

```ts
// outlook-addin/src/taskpane/useOutlookAgent.ts
import { useChat } from '@ai-sdk/react';
import { DefaultChatTransport, lastAssistantMessageIsCompleteWithToolCalls } from 'ai';

// Helper: get Graph access token via MSAL NAA
async function getGraphToken(): Promise<string> {
  // Use @azure/msal-browser with NestedAppAuthController
  // See: https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA
  throw new Error('Implement MSAL NAA token acquisition here');
}

const officeToolHandlers: Record<string, (args: Record<string, unknown>) => Promise<unknown>> = {

  getSelectedEmailBody: async () =>
    new Promise((resolve, reject) => {
      Office.context.mailbox.item?.body.getAsync(Office.CoercionType.Text, result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve({ body: result.value });
        } else {
          reject(new Error(result.error.message));
        }
      });
    }),

  getEmailMetadata: async () => {
    const item = Office.context.mailbox.item as Office.MessageRead;
    return {
      subject: item.subject,
      from: item.from?.emailAddress,
      to: item.to?.map(r => r.emailAddress),
      dateTimeCreated: item.dateTimeCreated?.toISOString(),
      itemId: item.itemId,
    };
  },

  flagEmail: async ({ flagged }) => {
    const item = Office.context.mailbox.item as Office.MessageRead;
    const restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
    const token = await getGraphToken();
    const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ flag: { flagStatus: flagged ? 'flagged' : 'notFlagged' } }),
    });
    if (!res.ok) throw new Error(`Graph error: ${res.status}`);
    return { success: true, flagged };
  },

  markAsRead: async ({ read }) => {
    const item = Office.context.mailbox.item as Office.MessageRead;
    const restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
    const token = await getGraphToken();
    const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ isRead: read }),
    });
    if (!res.ok) throw new Error(`Graph error: ${res.status}`);
    return { success: true, read };
  },

  openReplyForm: async ({ body }) => {
    Office.context.mailbox.item?.displayReplyAllForm({ htmlBody: `<p>${body}</p>` });
    return { success: true, note: 'Reply form opened. User must send manually.' };
  },

  showNotification: async ({ message, type }) => {
    Office.context.mailbox.item?.notificationMessages.addAsync(`notif-${Date.now()}`, {
      type: type as Office.MailboxEnums.ItemNotificationMessageType,
      message: message as string,
      icon: 'icon16',
      persistent: false,
    });
    return { shown: true, message };
  },
};

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
        // No await - avoids deadlock
        chat.addToolOutput({
          tool: toolCall.toolName,
          toolCallId: toolCall.toolCallId,
          state: 'output-error',
          errorText: `No Office.js handler for tool: ${toolCall.toolName}`,
        });
        return;
      }

      try {
        const output = await handler(toolCall.input as Record<string, unknown>);
        // No await - avoids deadlock
        chat.addToolOutput({
          tool: toolCall.toolName,
          toolCallId: toolCall.toolCallId,
          output,
        });
      } catch (err) {
        chat.addToolOutput({
          tool: toolCall.toolName,
          toolCallId: toolCall.toolCallId,
          state: 'output-error',
          errorText: err instanceof Error ? err.message : String(err),
        });
      }
    },
  });

  return chat;
}
```

---

## End-to-End Data Flow

```
User: "Flag this email and summarise it"
        |
        v
Add-in POST /api/chat  →  Agent App
        |
        v
streamText LLM decides to call:
  getEmailMetadata()     (no execute, forwarded to client)
  getSelectedEmailBody() (no execute, forwarded to client)
        |
        v
Stream arrives at Add-in browser
onToolCall fires for each call:
  item.subject / item.from  (sync)
  item.body.getAsync()      (async callback)
  addToolOutput() called for each
        |
        v
sendAutomaticallyWhen triggers re-POST to Agent App
        |
        v
streamText resumes, LLM calls:
  flagEmail({ flagged: true })
        |
        v
onToolCall fires:
  Graph PATCH /me/messages/{restId}
  addToolOutput({ success: true })
        |
        v
Agent streams final text reply to user
```

---

## Key Dependencies

Agent App:
- ai: ^5.x
- @ai-sdk/openai: ^1.x
- zod: ^3.x

Outlook Add-in:
- ai: ^5.x
- @ai-sdk/react: ^1.x
- @microsoft/office-js: latest
- @azure/msal-browser: ^3.x

---

## References

- Vercel AI SDK docs: https://sdk.vercel.ai/docs/ai-sdk-ui/chatbot-tool-usage
- execute-tool-call.ts source: https://github.com/vercel/ai/blob/main/packages/ai/src/generate-text/execute-tool-call.ts
- stream-text.ts source: https://github.com/vercel/ai/blob/main/packages/ai/src/generate-text/stream-text.ts
- OfficeDev Add-in Samples: https://github.com/OfficeDev/Office-Add-in-samples
- Outlook Add-in SSO NAA Sample: https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA
- Microsoft Graph update message: https://learn.microsoft.com/en-us/graph/api/message-update
