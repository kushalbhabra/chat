/**
 * Tests for TeamsSDKAdapter (index.ts)
 *
 * Tests the main adapter, TeamsApp event routing, and all adapter methods.
 */
import { describe, expect, it, vi, beforeEach } from "vitest";
import { parseMarkdown } from "chat";
import { TeamsApp, TeamsSDKAdapter, createTeamsSDKAdapter } from "./index.js";
import type { Activity } from "botbuilder";
import type { TurnContext } from "botbuilder";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeActivity(overrides: Partial<Activity> = {}): Activity {
  return {
    type: "message",
    id: "act-1",
    serviceUrl: "https://smba.trafficmanager.net/teams/",
    channelId: "msteams",
    from: { id: "user-1", name: "Alice" },
    conversation: { id: "conv-1", isGroup: true, tenantId: "tenant-1" },
    recipient: { id: "bot-1", name: "Bot" },
    text: "Hello",
    timestamp: new Date().toISOString(),
    ...overrides,
  } as Activity;
}

function makeTurnContext(activity: Activity): TurnContext {
  return {
    activity,
    sendActivity: vi.fn().mockResolvedValue({ id: "msg-1" }),
    updateActivity: vi.fn().mockResolvedValue(undefined),
    deleteActivity: vi.fn().mockResolvedValue(undefined),
  } as unknown as TurnContext;
}

// ---------------------------------------------------------------------------
// TeamsApp tests
// ---------------------------------------------------------------------------

describe("TeamsApp", () => {
  describe("$onBotActivity", () => {
    it("fires for every activity", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onBotActivity(handler);

      const activity = makeActivity();
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
      expect(handler).toHaveBeenCalledWith(activity, ctx);
    });
  });

  describe("$onMessage", () => {
    it("fires for regular group message (no replyToId, not DM, not mention)", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onMessage(handler);

      const activity = makeActivity({
        type: "message",
        conversation: { id: "conv-1", isGroup: true },
        entities: [],
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onMention", () => {
    it("fires when activity has a mention entity", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onMention(handler);

      const activity = makeActivity({
        type: "message",
        conversation: { id: "conv-1", isGroup: true },
        entities: [
          {
            type: "mention",
            mentioned: { id: "bot-1" },
          },
        ] as any,
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onThreadReplyAdded", () => {
    it("fires when activity has replyToId (non-DM)", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onThreadReplyAdded(handler);

      const activity = makeActivity({
        type: "message",
        replyToId: "parent-msg-1",
        conversation: { id: "conv-1", isGroup: true },
        entities: [],
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onDMReceived", () => {
    it("fires for personal/DM conversation", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onDMReceived(handler);

      const activity = makeActivity({
        type: "message",
        conversation: { id: "conv-1", isGroup: false, conversationType: "personal" },
        entities: [],
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onReactionAdded", () => {
    it("fires when reactionsAdded is non-empty", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onReactionAdded(handler);

      const activity = makeActivity({
        type: "messageReaction",
        reactionsAdded: [{ type: "like" }],
        reactionsRemoved: [],
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onReactionRemoved", () => {
    it("fires when reactionsRemoved is non-empty", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onReactionRemoved(handler);

      const activity = makeActivity({
        type: "messageReaction",
        reactionsAdded: [],
        reactionsRemoved: [{ type: "like" }],
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onCardAction", () => {
    it("fires for Action.Submit (message with value.actionId)", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onCardAction(handler);

      const activity = makeActivity({
        type: "message",
        value: { actionId: "my-action", value: "my-value" },
        text: "",
        conversation: { id: "conv-1", isGroup: true },
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });

    it("fires for adaptiveCard/action invoke", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onCardAction(handler);

      const activity = makeActivity({
        type: "invoke",
        name: "adaptiveCard/action",
        value: { action: { data: { actionId: "my-action" } } },
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onInvoke", () => {
    it("fires for invoke activities that are not adaptiveCard/action", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onInvoke(handler);

      const activity = makeActivity({
        type: "invoke",
        name: "task/fetch",
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onMemberAdded", () => {
    it("fires when conversationUpdate has membersAdded", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onMemberAdded(handler);

      const activity = makeActivity({
        type: "conversationUpdate",
        membersAdded: [{ id: "new-user" }],
        membersRemoved: [],
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onMemberRemoved", () => {
    it("fires when conversationUpdate has membersRemoved", async () => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app.$onMemberRemoved(handler);

      const activity = makeActivity({
        type: "conversationUpdate",
        membersAdded: [],
        membersRemoved: [{ id: "old-user" }],
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("channel events", () => {
    it.each([
      ["$onTeamRenamed", "teamRenamed"],
      ["$onChannelCreated", "channelCreated"],
      ["$onChannelRenamed", "channelRenamed"],
      ["$onChannelDeleted", "channelDeleted"],
    ] as const)("%s fires for eventType=%s", async (method, eventType) => {
      const app = new TeamsApp();
      const handler = vi.fn();
      app[method](handler);

      const activity = makeActivity({
        type: "conversationUpdate",
        membersAdded: [],
        membersRemoved: [],
        channelData: { eventType },
      });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(handler).toHaveBeenCalledOnce();
    });
  });

  describe("$onAppInstalled / $onAppUninstalled", () => {
    it("fires $onAppInstalled for installationUpdate add", async () => {
      const app = new TeamsApp();
      const installed = vi.fn();
      const uninstalled = vi.fn();
      app.$onAppInstalled(installed);
      app.$onAppUninstalled(uninstalled);

      const activity = makeActivity({ type: "installationUpdate", action: "add" });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(installed).toHaveBeenCalledOnce();
      expect(uninstalled).not.toHaveBeenCalled();
    });

    it("fires $onAppUninstalled for installationUpdate remove", async () => {
      const app = new TeamsApp();
      const installed = vi.fn();
      const uninstalled = vi.fn();
      app.$onAppInstalled(installed);
      app.$onAppUninstalled(uninstalled);

      const activity = makeActivity({ type: "installationUpdate", action: "remove" });
      const ctx = makeTurnContext(activity);
      await app.processActivity(activity, ctx);

      expect(uninstalled).toHaveBeenCalledOnce();
      expect(installed).not.toHaveBeenCalled();
    });
  });

  it("supports chaining registration methods", () => {
    const app = new TeamsApp();
    const result = app
      .$onMessage(vi.fn())
      .$onMention(vi.fn())
      .$onDMReceived(vi.fn())
      .$onReactionAdded(vi.fn());
    expect(result).toBe(app);
  });

  it("supports multiple handlers for the same event", async () => {
    const app = new TeamsApp();
    const h1 = vi.fn();
    const h2 = vi.fn();
    app.$onMessage(h1);
    app.$onMessage(h2);

    const activity = makeActivity({ type: "message", conversation: { id: "c", isGroup: true } });
    await app.processActivity(activity, makeTurnContext(activity));

    expect(h1).toHaveBeenCalledOnce();
    expect(h2).toHaveBeenCalledOnce();
  });
});

// ---------------------------------------------------------------------------
// TeamsSDKAdapter constructor tests
// ---------------------------------------------------------------------------

describe("TeamsSDKAdapter constructor", () => {
  it("throws when no appId provided", () => {
    const old = process.env.TEAMS_APP_ID;
    delete process.env.TEAMS_APP_ID;
    try {
      expect(() => new TeamsSDKAdapter({ appPassword: "pw" })).toThrow();
    } finally {
      if (old !== undefined) process.env.TEAMS_APP_ID = old;
    }
  });

  it("throws when no auth method provided", () => {
    expect(() => new TeamsSDKAdapter({ appId: "app-id" })).toThrow(
      /appPassword|certificate|federated/
    );
  });

  it("throws when multiple auth methods provided", () => {
    expect(() =>
      new TeamsSDKAdapter({
        appId: "app-id",
        appPassword: "pw",
        certificate: {
          certificatePrivateKey: "key",
          certificateThumbprint: "thumb",
        },
      })
    ).toThrow(/Only one/);
  });

  it("throws for SingleTenant without tenantId", () => {
    expect(() =>
      new TeamsSDKAdapter({
        appId: "app-id",
        appPassword: "pw",
        appType: "SingleTenant",
      })
    ).toThrow(/appTenantId/);
  });

  it("creates adapter successfully with appPassword", () => {
    const adapter = new TeamsSDKAdapter({
      appId: "app-id",
      appPassword: "pw",
    });
    expect(adapter.name).toBe("teamssdk");
    expect(adapter.userName).toBe("bot");
    expect(adapter.app).toBeInstanceOf(TeamsApp);
  });

  it("uses custom userName when provided", () => {
    const adapter = new TeamsSDKAdapter({
      appId: "app-id",
      appPassword: "pw",
      userName: "my-bot",
    });
    expect(adapter.userName).toBe("my-bot");
  });
});

// ---------------------------------------------------------------------------
// Thread ID encoding/decoding
// ---------------------------------------------------------------------------

describe("encodeThreadId / decodeThreadId", () => {
  const adapter = new TeamsSDKAdapter({ appId: "app-id", appPassword: "pw" });

  it("encodes and decodes a round-trip", () => {
    const data = {
      conversationId: "19:abc@thread.tacv2",
      serviceUrl: "https://smba.trafficmanager.net/teams/",
    };
    const encoded = adapter.encodeThreadId(data);
    expect(encoded.startsWith("teamssdk:")).toBe(true);
    const decoded = adapter.decodeThreadId(encoded);
    expect(decoded.conversationId).toBe(data.conversationId);
    expect(decoded.serviceUrl).toBe(data.serviceUrl);
  });

  it("encodes replyToId in thread ID", () => {
    const data = {
      conversationId: "19:abc@thread.tacv2",
      serviceUrl: "https://smba.trafficmanager.net/",
      replyToId: "msg-123",
    };
    const encoded = adapter.encodeThreadId(data);
    const decoded = adapter.decodeThreadId(encoded);
    expect(decoded.replyToId).toBe("msg-123");
  });

  it("falls back gracefully for non-base64 input", () => {
    const decoded = adapter.decodeThreadId("not-valid-base64!@#$");
    expect(decoded).toBeDefined();
  });
});

describe("isDM", () => {
  const adapter = new TeamsSDKAdapter({ appId: "app-id", appPassword: "pw" });

  it("returns false for channel thread ID", () => {
    const threadId = adapter.encodeThreadId({
      conversationId: "19:abc@thread.tacv2",
      serviceUrl: "https://smba.trafficmanager.net/",
    });
    expect(adapter.isDM(threadId)).toBe(false);
  });

  it("returns true for personal conversation ID", () => {
    const threadId = adapter.encodeThreadId({
      conversationId: "a:AbCdEf1234567890",
      serviceUrl: "https://smba.trafficmanager.net/",
    });
    expect(adapter.isDM(threadId)).toBe(true);
  });
});

describe("channelIdFromThreadId", () => {
  const adapter = new TeamsSDKAdapter({ appId: "app-id", appPassword: "pw" });

  it("strips ;messageid= suffix", () => {
    const threadId = adapter.encodeThreadId({
      conversationId: "19:abc@thread.tacv2;messageid=12345",
      serviceUrl: "https://smba.trafficmanager.net/",
    });
    const channelId = adapter.channelIdFromThreadId(threadId);
    expect(channelId).toBe("19:abc@thread.tacv2");
    expect(channelId).not.toContain("messageid");
  });
});

// ---------------------------------------------------------------------------
// parseMessage
// ---------------------------------------------------------------------------

describe("parseMessage", () => {
  const adapter = new TeamsSDKAdapter({ appId: "app-id", appPassword: "pw" });

  it("converts a Teams activity to a normalized Message", () => {
    const activity = makeActivity({ text: "Hello **world**" });
    const message = adapter.parseMessage(activity);
    expect(message.id).toBe("act-1");
    expect(message.text).toBeTruthy();
    expect(message.formatted).toBeTruthy();
    expect(message.raw).toBe(activity);
  });

  it("sets author fields from activity.from", () => {
    const activity = makeActivity({
      from: { id: "u-123", name: "Bob" },
    });
    const message = adapter.parseMessage(activity);
    expect(message.author.userId).toBe("u-123");
    expect(message.author.userName).toBe("Bob");
  });
});

// ---------------------------------------------------------------------------
// renderFormatted
// ---------------------------------------------------------------------------

describe("renderFormatted", () => {
  const adapter = new TeamsSDKAdapter({ appId: "app-id", appPassword: "pw" });

  it("renders an AST to a Teams-formatted string", () => {
    const ast = parseMarkdown("**bold** and _italic_");
    const result = adapter.renderFormatted(ast);
    expect(result).toContain("bold");
    expect(result).toContain("italic");
  });
});

// ---------------------------------------------------------------------------
// createTeamsSDKAdapter factory
// ---------------------------------------------------------------------------

describe("createTeamsSDKAdapter", () => {
  it("creates an instance with factory function", () => {
    const adapter = createTeamsSDKAdapter({ appId: "app-id", appPassword: "pw" });
    expect(adapter).toBeInstanceOf(TeamsSDKAdapter);
  });
});
