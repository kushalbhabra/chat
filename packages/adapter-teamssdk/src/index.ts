/**
 * Teams SDK Adapter for vercel/chat
 *
 * Implements the Adapter<TeamsThreadId, unknown> interface using botbuilder,
 * wrapped in a teams.ts-style event-routing abstraction (TeamsApp).
 *
 * Architecture:
 *   handleWebhook → botAdapter.processActivity → TeamsApp.processActivity
 *     → $onMessage / $onMention / $onCardAction / $onReactionAdded / …
 *       → chat.processMessage / chat.processAction / chat.processReaction
 */

import type { TokenCredential } from "@azure/identity";
import {
  ClientCertificateCredential,
  ClientSecretCredential,
  DefaultAzureCredential,
} from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import {
  TokenCredentialAuthenticationProvider,
  type TokenCredentialAuthenticationProviderOptions,
} from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import type { Activity, ConversationReference } from "botbuilder";
import {
  ActivityTypes,
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TeamsInfo,
  type TurnContext,
} from "botbuilder";
import {
  CertificateServiceClientCredentialsFactory,
  FederatedServiceClientCredentialsFactory,
} from "botframework-connector";

import {
  AdapterRateLimitError,
  AuthenticationError,
  bufferToDataUri,
  extractCard,
  extractFiles,
  NetworkError,
  PermissionError,
  toBuffer,
  ValidationError,
} from "@chat-adapter/shared";
import type {
  ActionEvent,
  Adapter,
  AdapterPostableMessage,
  Attachment,
  ChannelInfo,
  ChatInstance,
  EmojiValue,
  FetchOptions,
  FetchResult,
  FileUpload,
  FormattedContent,
  ListThreadsOptions,
  ListThreadsResult,
  Logger,
  RawMessage,
  ReactionEvent,
  ThreadInfo,
  WebhookOptions,
} from "chat";
import {
  ConsoleLogger,
  convertEmojiPlaceholders,
  defaultEmojiResolver,
  Message,
  NotImplementedError,
} from "chat";
import { cardToAdaptiveCard } from "./cards.js";
import { TeamsSDKFormatConverter } from "./markdown.js";

const MESSAGEID_CAPTURE_PATTERN = /messageid=(\d+)/;
const MESSAGEID_STRIP_PATTERN = /;messageid=\d+/;
const SEMICOLON_MESSAGEID_CAPTURE_PATTERN = /;messageid=(\d+)/;

// ---------------------------------------------------------------------------
// TeamsApp – the teams.ts-style event-routing abstraction
// ---------------------------------------------------------------------------

/** Handler function type for Teams events */
export type TeamsEventHandler = (
  activity: Activity,
  context: TurnContext
) => Promise<void> | void;

/**
 * TeamsApp provides an event-driven API for routing Teams Bot Framework
 * activities to typed handlers. This mirrors the teams.ts SDK pattern,
 * where each event type has its own registration method ($onMessage, etc.).
 *
 * @example
 * ```ts
 * const app = new TeamsApp();
 * app.$onMessage(async (activity) => {
 *   console.log("Received message:", activity.text);
 * });
 * ```
 */
export class TeamsApp {
  private readonly handlers: Map<string, TeamsEventHandler[]> = new Map();

  // -- Registration methods --

  /** Handle regular channel/group messages */
  $onMessage(handler: TeamsEventHandler): this {
    return this.addHandler("message", handler);
  }

  /** Handle @mention messages (the bot was explicitly mentioned) */
  $onMention(handler: TeamsEventHandler): this {
    return this.addHandler("mention", handler);
  }

  /** Handle replies in a thread */
  $onThreadReplyAdded(handler: TeamsEventHandler): this {
    return this.addHandler("threadReply", handler);
  }

  /** Handle 1:1 direct messages sent to the bot */
  $onDMReceived(handler: TeamsEventHandler): this {
    return this.addHandler("dm", handler);
  }

  /** Handle emoji reactions added to a message */
  $onReactionAdded(handler: TeamsEventHandler): this {
    return this.addHandler("reactionAdded", handler);
  }

  /** Handle emoji reactions removed from a message */
  $onReactionRemoved(handler: TeamsEventHandler): this {
    return this.addHandler("reactionRemoved", handler);
  }

  /** Handle Adaptive Card Action.Submit clicks */
  $onCardAction(handler: TeamsEventHandler): this {
    return this.addHandler("cardAction", handler);
  }

  /** Handle invoke activities (task modules, adaptive card invokes) */
  $onInvoke(handler: TeamsEventHandler): this {
    return this.addHandler("invoke", handler);
  }

  /** Handle message context-menu actions */
  $onMessageAction(handler: TeamsEventHandler): this {
    return this.addHandler("messageAction", handler);
  }

  /** Handle member added to team/conversation */
  $onMemberAdded(handler: TeamsEventHandler): this {
    return this.addHandler("memberAdded", handler);
  }

  /** Handle member removed from team/conversation */
  $onMemberRemoved(handler: TeamsEventHandler): this {
    return this.addHandler("memberRemoved", handler);
  }

  /** Handle team rename event */
  $onTeamRenamed(handler: TeamsEventHandler): this {
    return this.addHandler("teamRenamed", handler);
  }

  /** Handle channel created in a team */
  $onChannelCreated(handler: TeamsEventHandler): this {
    return this.addHandler("channelCreated", handler);
  }

  /** Handle channel renamed in a team */
  $onChannelRenamed(handler: TeamsEventHandler): this {
    return this.addHandler("channelRenamed", handler);
  }

  /** Handle channel deleted from a team */
  $onChannelDeleted(handler: TeamsEventHandler): this {
    return this.addHandler("channelDeleted", handler);
  }

  /** Handle app installation in a team/personal scope */
  $onAppInstalled(handler: TeamsEventHandler): this {
    return this.addHandler("appInstalled", handler);
  }

  /** Handle app uninstallation */
  $onAppUninstalled(handler: TeamsEventHandler): this {
    return this.addHandler("appUninstalled", handler);
  }

  /** Handle any bot activity (fires for every activity, regardless of type) */
  $onBotActivity(handler: TeamsEventHandler): this {
    return this.addHandler("botActivity", handler);
  }

  // -- Dispatch --

  /**
   * Route an incoming activity to all matching registered handlers.
   * Always fires $onBotActivity handlers, then type-specific handlers.
   */
  async processActivity(
    activity: Activity,
    context: TurnContext
  ): Promise<void> {
    // Always fire generic activity handlers first
    await this.runHandlers("botActivity", activity, context);

    switch (activity.type) {
      case ActivityTypes.Message:
        await this.dispatchMessageActivity(activity, context);
        break;

      case ActivityTypes.MessageReaction:
        await this.dispatchReactionActivity(activity, context);
        break;

      case ActivityTypes.Invoke:
        await this.dispatchInvokeActivity(activity, context);
        break;

      case ActivityTypes.ConversationUpdate:
        await this.dispatchConversationUpdateActivity(activity, context);
        break;

      case ActivityTypes.InstallationUpdate:
        await this.dispatchInstallationUpdateActivity(activity, context);
        break;
    }
  }

  // -- Private dispatch helpers --

  private async dispatchMessageActivity(
    activity: Activity,
    context: TurnContext
  ): Promise<void> {
    // Action.Submit sends as message with value.actionId
    const actionValue = activity.value as
      | { actionId?: string }
      | undefined;
    if (actionValue?.actionId) {
      await this.runHandlers("cardAction", activity, context);
      return;
    }

    const isDM =
      activity.conversation?.conversationType === "personal" ||
      !activity.conversation?.isGroup;
    const isMention = (activity.entities ?? []).some(
      (e) => e.type === "mention" && (e as { mentioned?: { id?: string } }).mentioned?.id
    );
    const isReply = !!activity.replyToId;

    if (isDM) {
      await this.runHandlers("dm", activity, context);
    } else if (isMention) {
      await this.runHandlers("mention", activity, context);
    } else if (isReply) {
      await this.runHandlers("threadReply", activity, context);
    } else {
      await this.runHandlers("message", activity, context);
    }
  }

  private async dispatchReactionActivity(
    activity: Activity,
    context: TurnContext
  ): Promise<void> {
    if ((activity.reactionsAdded ?? []).length > 0) {
      await this.runHandlers("reactionAdded", activity, context);
    }
    if ((activity.reactionsRemoved ?? []).length > 0) {
      await this.runHandlers("reactionRemoved", activity, context);
    }
  }

  private async dispatchInvokeActivity(
    activity: Activity,
    context: TurnContext
  ): Promise<void> {
    if (activity.name === "adaptiveCard/action") {
      await this.runHandlers("cardAction", activity, context);
    } else {
      await this.runHandlers("invoke", activity, context);
    }
  }

  private async dispatchConversationUpdateActivity(
    activity: Activity,
    context: TurnContext
  ): Promise<void> {
    const channelData = activity.channelData as Record<string, unknown> | undefined;
    const eventType = channelData?.eventType as string | undefined;

    if ((activity.membersAdded ?? []).length > 0) {
      await this.runHandlers("memberAdded", activity, context);
    }
    if ((activity.membersRemoved ?? []).length > 0) {
      await this.runHandlers("memberRemoved", activity, context);
    }

    if (eventType === "teamRenamed") {
      await this.runHandlers("teamRenamed", activity, context);
    } else if (eventType === "channelCreated") {
      await this.runHandlers("channelCreated", activity, context);
    } else if (eventType === "channelRenamed") {
      await this.runHandlers("channelRenamed", activity, context);
    } else if (eventType === "channelDeleted") {
      await this.runHandlers("channelDeleted", activity, context);
    }
  }

  private async dispatchInstallationUpdateActivity(
    activity: Activity,
    context: TurnContext
  ): Promise<void> {
    if (activity.action === "add") {
      await this.runHandlers("appInstalled", activity, context);
    } else if (activity.action === "remove") {
      await this.runHandlers("appUninstalled", activity, context);
    }
  }

  private addHandler(event: string, handler: TeamsEventHandler): this {
    const existing = this.handlers.get(event) ?? [];
    existing.push(handler);
    this.handlers.set(event, existing);
    return this;
  }

  private async runHandlers(
    event: string,
    activity: Activity,
    context: TurnContext
  ): Promise<void> {
    const list = this.handlers.get(event) ?? [];
    for (const handler of list) {
      await handler(activity, context);
    }
  }
}

// ---------------------------------------------------------------------------
// Microsoft Graph chat message type
// ---------------------------------------------------------------------------

interface GraphChatMessage {
  attachments?: Array<{
    id?: string;
    contentType?: string;
    contentUrl?: string;
    content?: string;
    name?: string;
  }>;
  body?: {
    content?: string;
    contentType?: "text" | "html";
  };
  createdDateTime?: string;
  from?: {
    user?: {
      id?: string;
      displayName?: string;
    };
    application?: {
      id?: string;
      displayName?: string;
    };
  };
  id: string;
  lastModifiedDateTime?: string;
  replyToId?: string;
}

// ---------------------------------------------------------------------------
// Auth config types
// ---------------------------------------------------------------------------

/** Certificate-based authentication config */
export interface TeamsAuthCertificate {
  /** PEM-encoded certificate private key */
  certificatePrivateKey: string;
  /** Hex-encoded certificate thumbprint (optional when x5c is provided) */
  certificateThumbprint?: string;
  /** Public certificate for subject-name validation (optional) */
  x5c?: string;
}

/** Federated (workload identity) authentication config */
export interface TeamsAuthFederated {
  /** Audience for the federated credential (defaults to api://AzureADTokenExchange) */
  clientAudience?: string;
  /** Client ID for the managed identity assigned to the bot */
  clientId: string;
}

export interface TeamsSDKAdapterConfig {
  /** Microsoft App ID. Defaults to TEAMS_APP_ID env var. */
  appId?: string;
  /** Microsoft App Password. Defaults to TEAMS_APP_PASSWORD env var. */
  appPassword?: string;
  /** Microsoft App Tenant ID. Defaults to TEAMS_APP_TENANT_ID env var. */
  appTenantId?: string;
  /** Microsoft App Type */
  appType?: "MultiTenant" | "SingleTenant";
  /** Certificate-based authentication */
  certificate?: TeamsAuthCertificate;
  /** Federated (workload identity) authentication */
  federated?: TeamsAuthFederated;
  /** Logger instance. Defaults to ConsoleLogger. */
  logger?: Logger;
  /** Override bot username (optional) */
  userName?: string;
}

/** Teams-specific thread ID data */
export interface TeamsThreadId {
  conversationId: string;
  replyToId?: string;
  serviceUrl: string;
}

/** Teams channel context extracted from activity.channelData */
interface TeamsChannelContext {
  channelId: string;
  teamId: string;
  tenantId: string;
}

// ---------------------------------------------------------------------------
// Extend CloudAdapter for serverless (no HTTP server needed)
// ---------------------------------------------------------------------------

class ServerlessCloudAdapter extends CloudAdapter {
  handleActivity(
    authHeader: string,
    activity: Activity,
    logic: (context: TurnContext) => Promise<void>
  ) {
    return this.processActivity(authHeader, activity, logic);
  }
}

// ---------------------------------------------------------------------------
// TeamsSDKAdapter – main adapter class
// ---------------------------------------------------------------------------

export class TeamsSDKAdapter implements Adapter<TeamsThreadId, unknown> {
  readonly name = "teamssdk";
  readonly userName: string;
  readonly botUserId?: string;

  /** The internal TeamsApp event router (teams.ts-style API) */
  readonly app: TeamsApp;

  private readonly botAdapter: ServerlessCloudAdapter;
  private readonly graphClient: Client | null = null;
  private chat: ChatInstance | null = null;
  private readonly logger: Logger;
  private readonly formatConverter = new TeamsSDKFormatConverter();
  private readonly config: Required<Pick<TeamsSDKAdapterConfig, "appId">> &
    TeamsSDKAdapterConfig;

  constructor(config: TeamsSDKAdapterConfig = {}) {
    const appId = config.appId ?? process.env.TEAMS_APP_ID;
    if (!appId) {
      throw new ValidationError(
        "teamssdk",
        "appId is required. Set TEAMS_APP_ID or provide it in config."
      );
    }

    const hasExplicitAuth =
      config.appPassword || config.certificate || config.federated;
    const appPassword = hasExplicitAuth
      ? config.appPassword
      : (config.appPassword ?? process.env.TEAMS_APP_PASSWORD);
    const appTenantId = config.appTenantId ?? process.env.TEAMS_APP_TENANT_ID;

    this.config = { ...config, appId, appPassword, appTenantId };
    this.logger = config.logger ?? new ConsoleLogger("info").child("teamssdk");
    this.userName = config.userName || "bot";

    // Validate auth config
    const authMethodCount = [
      appPassword,
      config.certificate,
      config.federated,
    ].filter(Boolean).length;

    if (authMethodCount === 0) {
      throw new ValidationError(
        "teamssdk",
        "One of appPassword, certificate, or federated must be provided"
      );
    }
    if (authMethodCount > 1) {
      throw new ValidationError(
        "teamssdk",
        "Only one of appPassword, certificate, or federated can be provided"
      );
    }
    if (config.appType === "SingleTenant" && !appTenantId) {
      throw new ValidationError(
        "teamssdk",
        "appTenantId is required for SingleTenant app type"
      );
    }

    const botFrameworkConfig = {
      MicrosoftAppId: appId,
      MicrosoftAppType: config.appType || "MultiTenant",
      MicrosoftAppTenantId:
        config.appType === "SingleTenant" ? appTenantId : undefined,
    };

    let credentialsFactory:
      | CertificateServiceClientCredentialsFactory
      | FederatedServiceClientCredentialsFactory
      | undefined;
    let graphCredential: TokenCredential | undefined;

    if (config.certificate) {
      const { certificatePrivateKey, certificateThumbprint, x5c } =
        config.certificate;
      if (x5c) {
        credentialsFactory = new CertificateServiceClientCredentialsFactory(
          appId,
          x5c,
          certificatePrivateKey,
          appTenantId
        );
      } else if (certificateThumbprint) {
        credentialsFactory = new CertificateServiceClientCredentialsFactory(
          appId,
          certificateThumbprint,
          certificatePrivateKey,
          appTenantId
        );
      } else {
        throw new ValidationError(
          "teamssdk",
          "Certificate auth requires either certificateThumbprint or x5c"
        );
      }
      if (appTenantId) {
        graphCredential = new ClientCertificateCredential(appTenantId, appId, {
          certificate: certificatePrivateKey,
        });
      }
    } else if (config.federated) {
      credentialsFactory = new FederatedServiceClientCredentialsFactory(
        appId,
        config.federated.clientId,
        appTenantId,
        config.federated.clientAudience
      );
      if (appTenantId) {
        graphCredential = new DefaultAzureCredential();
      }
    } else if (appPassword && appTenantId) {
      graphCredential = new ClientSecretCredential(
        appTenantId,
        appId,
        appPassword
      );
    }

    const auth = new ConfigurationBotFrameworkAuthentication(
      {
        ...botFrameworkConfig,
        ...(appPassword ? { MicrosoftAppPassword: appPassword } : {}),
      },
      credentialsFactory
    );

    this.botAdapter = new ServerlessCloudAdapter(auth);

    if (graphCredential) {
      const authProvider = new TokenCredentialAuthenticationProvider(
        graphCredential,
        {
          scopes: ["https://graph.microsoft.com/.default"],
        } as TokenCredentialAuthenticationProviderOptions
      );
      this.graphClient = Client.initWithMiddleware({ authProvider });
    }

    // -- Build the TeamsApp event router and wire it to chat --
    this.app = new TeamsApp();
    this.wireAppHandlers();
  }

  /**
   * Wire up the TeamsApp event handlers to forward activities to the
   * Chat instance (processMessage / processAction / processReaction).
   */
  private wireAppHandlers(): void {
    // Message events
    for (const event of ["message", "mention", "threadReply", "dm"] as const) {
      this.app[
        event === "message"
          ? "$onMessage"
          : event === "mention"
            ? "$onMention"
            : event === "threadReply"
              ? "$onThreadReplyAdded"
              : "$onDMReceived"
      ]((activity) => this.handleMessageActivity(activity));
    }

    // Reaction events
    this.app.$onReactionAdded((activity) =>
      this.handleReactionActivity(activity, true)
    );
    this.app.$onReactionRemoved((activity) =>
      this.handleReactionActivity(activity, false)
    );

    // Card action events (invokes and message actions)
    this.app.$onCardAction((activity, context) =>
      this.handleCardActionActivity(activity, context)
    );
    this.app.$onInvoke((activity, context) =>
      this.handleInvokeActivity(activity, context)
    );
  }

  async initialize(chat: ChatInstance): Promise<void> {
    this.chat = chat;
  }

  async handleWebhook(
    request: Request,
    options?: WebhookOptions
  ): Promise<Response> {
    const body = await request.text();
    this.logger.debug("Teams SDK webhook raw body", { body });

    let activity: Activity;
    try {
      activity = JSON.parse(body) as Activity;
    } catch (e) {
      this.logger.error("Failed to parse request body", { error: e });
      return new Response("Invalid JSON", { status: 400 });
    }

    const authHeader = request.headers.get("authorization") ?? "";

    try {
      await this.botAdapter.handleActivity(
        authHeader,
        activity,
        async (context) => {
          await this.handleTurn(context, options);
        }
      );

      return new Response(JSON.stringify({}), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    } catch (error) {
      this.logger.error("Bot adapter process error", { error });
      return new Response(JSON.stringify({ error: "Internal error" }), {
        status: 500,
        headers: { "Content-Type": "application/json" },
      });
    }
  }

  private async handleTurn(
    context: TurnContext,
    _options?: WebhookOptions
  ): Promise<void> {
    if (!this.chat) {
      this.logger.warn("Chat instance not initialized, ignoring event");
      return;
    }

    const activity = context.activity;

    // Cache serviceUrl and tenantId for future DM creation
    if (activity.from?.id && activity.serviceUrl) {
      const userId = activity.from.id;
      const channelData = activity.channelData as {
        tenant?: { id?: string };
        team?: { id?: string; aadGroupId?: string };
        channel?: { id?: string };
      };
      const tenantId = channelData?.tenant?.id;
      const ttl = 30 * 24 * 60 * 60 * 1000;

      this.chat
        .getState()
        .set(`teamssdk:serviceUrl:${userId}`, activity.serviceUrl, ttl)
        .catch((err) => {
          this.logger.error("Failed to cache serviceUrl", { userId, error: err });
        });

      if (tenantId) {
        this.chat
          .getState()
          .set(`teamssdk:tenantId:${userId}`, tenantId, ttl)
          .catch((err) => {
            this.logger.error("Failed to cache tenantId", { userId, error: err });
          });
      }

      // Cache team/channel context for Graph API message fetching
      const team = channelData?.team as
        | { id?: string; aadGroupId?: string }
        | undefined;
      const teamAadGroupId = team?.aadGroupId;
      const teamThreadId = team?.id;
      const conversationId = activity.conversation?.id ?? "";
      const baseChannelId = conversationId.replace(MESSAGEID_STRIP_PATTERN, "");

      if (teamAadGroupId && channelData?.channel?.id && tenantId) {
        const ctx: TeamsChannelContext = {
          teamId: teamAadGroupId,
          channelId: channelData.channel.id,
          tenantId,
        };
        const ctxJson = JSON.stringify(ctx);

        this.chat
          .getState()
          .set(`teamssdk:channelContext:${baseChannelId}`, ctxJson, ttl)
          .catch(() => undefined);

        if (teamThreadId) {
          this.chat
            .getState()
            .set(`teamssdk:teamContext:${teamThreadId}`, ctxJson, ttl)
            .catch(() => undefined);
        }
      } else if (teamThreadId && channelData?.channel?.id && tenantId) {
        const cachedTeamContext = await this.chat
          .getState()
          .get<string>(`teamssdk:teamContext:${teamThreadId}`);

        if (cachedTeamContext) {
          this.chat
            .getState()
            .set(`teamssdk:channelContext:${baseChannelId}`, cachedTeamContext, ttl)
            .catch(() => undefined);
        } else {
          try {
            const teamDetails = await TeamsInfo.getTeamDetails(context);
            if (teamDetails?.aadGroupId) {
              const fetchedCtx: TeamsChannelContext = {
                teamId: teamDetails.aadGroupId,
                channelId: channelData.channel.id,
                tenantId,
              };
              const fetchedJson = JSON.stringify(fetchedCtx);
              this.chat
                .getState()
                .set(`teamssdk:channelContext:${baseChannelId}`, fetchedJson, ttl)
                .catch(() => undefined);
              this.chat
                .getState()
                .set(`teamssdk:teamContext:${teamThreadId}`, fetchedJson, ttl)
                .catch(() => undefined);
            }
          } catch {
            // TeamsInfo.getTeamDetails() only works in team scope
          }
        }
      }
    }

    // Dispatch to TeamsApp router
    await this.app.processActivity(activity, context);
  }

  // ---------------------------------------------------------------------------
  // TeamsApp event implementation handlers
  // ---------------------------------------------------------------------------

  private handleMessageActivity(activity: Activity, options?: WebhookOptions): void {
    if (!this.chat) return;

    const threadId = this.encodeThreadId({
      conversationId: activity.conversation?.id ?? "",
      serviceUrl: activity.serviceUrl ?? "",
      replyToId: activity.replyToId,
    });

    this.chat.processMessage(
      this,
      threadId,
      this.parseTeamsMessage(activity, threadId),
      options
    );
  }

  private handleReactionActivity(
    activity: Activity,
    added: boolean,
    options?: WebhookOptions
  ): void {
    if (!this.chat) return;

    const conversationId = activity.conversation?.id ?? "";
    const messageIdMatch = conversationId.match(MESSAGEID_CAPTURE_PATTERN);
    const messageId = messageIdMatch?.[1] ?? activity.replyToId ?? "";

    const threadId = this.encodeThreadId({
      conversationId,
      serviceUrl: activity.serviceUrl ?? "",
    });

    const user = {
      userId: activity.from?.id ?? "unknown",
      userName: activity.from?.name ?? "unknown",
      fullName: activity.from?.name,
      isBot: false,
      isMe: this.isMessageFromSelf(activity),
    };

    const reactions = added
      ? (activity.reactionsAdded ?? [])
      : (activity.reactionsRemoved ?? []);

    for (const reaction of reactions) {
      const rawEmoji = reaction.type ?? "";
      const emojiValue = defaultEmojiResolver.fromTeams(rawEmoji);

      const event: Omit<ReactionEvent, "adapter" | "thread"> = {
        emoji: emojiValue,
        rawEmoji,
        added,
        user,
        messageId,
        threadId,
        raw: activity,
      };

      this.chat.processReaction({ ...event, adapter: this }, options);
    }
  }

  private handleCardActionActivity(
    activity: Activity,
    context: TurnContext,
    options?: WebhookOptions
  ): void {
    if (!this.chat) return;

    // Action.Submit (message type with value.actionId)
    const actionValue = activity.value as
      | { actionId?: string; value?: string }
      | undefined;
    if (!actionValue?.actionId) return;

    const threadId = this.encodeThreadId({
      conversationId: activity.conversation?.id ?? "",
      serviceUrl: activity.serviceUrl ?? "",
    });

    const actionEvent: Omit<ActionEvent, "thread" | "openModal"> & {
      adapter: TeamsSDKAdapter;
    } = {
      actionId: actionValue.actionId,
      value: actionValue.value,
      user: {
        userId: activity.from?.id ?? "unknown",
        userName: activity.from?.name ?? "unknown",
        fullName: activity.from?.name ?? "unknown",
        isBot: false,
        isMe: false,
      },
      messageId: activity.replyToId ?? activity.id ?? "",
      threadId,
      adapter: this,
      raw: activity,
    };

    this.chat.processAction(actionEvent, options);
  }

  private async handleInvokeActivity(
    activity: Activity,
    context: TurnContext,
    options?: WebhookOptions
  ): Promise<void> {
    if (!this.chat) return;

    if (activity.name === "adaptiveCard/action") {
      const actionData = (activity.value as { action?: { data?: { actionId?: string; value?: string } } })
        ?.action?.data;

      if (!actionData?.actionId) {
        await context.sendActivity({
          type: ActivityTypes.InvokeResponse,
          value: { status: 200 },
        });
        return;
      }

      const threadId = this.encodeThreadId({
        conversationId: activity.conversation?.id ?? "",
        serviceUrl: activity.serviceUrl ?? "",
      });

      const actionEvent: Omit<ActionEvent, "thread" | "openModal"> & {
        adapter: TeamsSDKAdapter;
      } = {
        actionId: actionData.actionId,
        value: actionData.value,
        user: {
          userId: activity.from?.id ?? "unknown",
          userName: activity.from?.name ?? "unknown",
          fullName: activity.from?.name ?? "unknown",
          isBot: false,
          isMe: false,
        },
        messageId: activity.replyToId ?? activity.id ?? "",
        threadId,
        adapter: this,
        raw: activity,
      };

      this.chat.processAction(actionEvent, options);

      await context.sendActivity({
        type: ActivityTypes.InvokeResponse,
        value: { status: 200 },
      });
    }
  }

  // ---------------------------------------------------------------------------
  // Adapter interface implementation
  // ---------------------------------------------------------------------------

  async postMessage(
    threadId: string,
    message: AdapterPostableMessage
  ): Promise<RawMessage<unknown>> {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);

    const files = extractFiles(message);
    const fileAttachments =
      files.length > 0 ? await this.filesToAttachments(files) : [];

    const card = extractCard(message);
    let activity: Partial<Activity>;

    if (card) {
      const adaptiveCard = cardToAdaptiveCard(card);
      activity = {
        type: ActivityTypes.Message,
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: adaptiveCard,
          },
          ...fileAttachments,
        ],
      };
    } else {
      const text = convertEmojiPlaceholders(
        this.formatConverter.renderPostable(message),
        "teams"
      );
      activity = {
        type: ActivityTypes.Message,
        text,
        textFormat: "markdown",
        attachments: fileAttachments.length > 0 ? fileAttachments : undefined,
      };
    }

    const conversationReference = {
      channelId: "msteams",
      serviceUrl,
      conversation: { id: conversationId },
    };

    let messageId = "";
    try {
      await this.botAdapter.continueConversationAsync(
        this.config.appId,
        conversationReference as Partial<ConversationReference>,
        async (context) => {
          const response = await context.sendActivity(activity);
          messageId = response?.id ?? "";
        }
      );
    } catch (error) {
      this.handleTeamsError(error, "postMessage");
    }

    return { id: messageId, threadId, raw: activity };
  }

  private async filesToAttachments(
    files: FileUpload[]
  ): Promise<Array<{ contentType: string; contentUrl: string; name: string }>> {
    const attachments: Array<{
      contentType: string;
      contentUrl: string;
      name: string;
    }> = [];

    for (const file of files) {
      const buffer = await toBuffer(file.data, {
        platform: "teams",
        throwOnUnsupported: false,
      });
      if (!buffer) continue;

      const mimeType = file.mimeType ?? "application/octet-stream";
      const dataUri = bufferToDataUri(buffer, mimeType);
      attachments.push({ contentType: mimeType, contentUrl: dataUri, name: file.filename });
    }

    return attachments;
  }

  async editMessage(
    threadId: string,
    messageId: string,
    message: AdapterPostableMessage
  ): Promise<RawMessage<unknown>> {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);

    const card = extractCard(message);
    let activity: Partial<Activity>;

    if (card) {
      const adaptiveCard = cardToAdaptiveCard(card);
      activity = {
        id: messageId,
        type: ActivityTypes.Message,
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: adaptiveCard,
          },
        ],
      };
    } else {
      const text = convertEmojiPlaceholders(
        this.formatConverter.renderPostable(message),
        "teams"
      );
      activity = { id: messageId, type: ActivityTypes.Message, text, textFormat: "markdown" };
    }

    const conversationReference = {
      channelId: "msteams",
      serviceUrl,
      conversation: { id: conversationId },
    };

    try {
      await this.botAdapter.continueConversationAsync(
        this.config.appId,
        conversationReference as Partial<ConversationReference>,
        async (context) => {
          await context.updateActivity(activity);
        }
      );
    } catch (error) {
      this.handleTeamsError(error, "editMessage");
    }

    return { id: messageId, threadId, raw: activity };
  }

  async deleteMessage(threadId: string, messageId: string): Promise<void> {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);

    const conversationReference = {
      channelId: "msteams",
      serviceUrl,
      conversation: { id: conversationId },
    };

    try {
      await this.botAdapter.continueConversationAsync(
        this.config.appId,
        conversationReference as Partial<ConversationReference>,
        async (context) => {
          await context.deleteActivity(messageId);
        }
      );
    } catch (error) {
      this.handleTeamsError(error, "deleteMessage");
    }
  }

  async addReaction(
    _threadId: string,
    _messageId: string,
    _emoji: EmojiValue | string
  ): Promise<void> {
    throw new NotImplementedError(
      "Teams Bot Framework does not expose reaction APIs",
      "addReaction"
    );
  }

  async removeReaction(
    _threadId: string,
    _messageId: string,
    _emoji: EmojiValue | string
  ): Promise<void> {
    throw new NotImplementedError(
      "Teams Bot Framework does not expose reaction APIs",
      "removeReaction"
    );
  }

  async startTyping(threadId: string, _status?: string): Promise<void> {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);

    const conversationReference = {
      channelId: "msteams",
      serviceUrl,
      conversation: { id: conversationId },
    };

    try {
      await this.botAdapter.continueConversationAsync(
        this.config.appId,
        conversationReference as Partial<ConversationReference>,
        async (context) => {
          await context.sendActivity({ type: ActivityTypes.Typing });
        }
      );
    } catch (error) {
      this.handleTeamsError(error, "startTyping");
    }
  }

  async openDM(userId: string): Promise<string> {
    const cachedServiceUrl = await this.chat
      ?.getState()
      .get<string>(`teamssdk:serviceUrl:${userId}`);
    const cachedTenantId = await this.chat
      ?.getState()
      .get<string>(`teamssdk:tenantId:${userId}`);

    const serviceUrl =
      cachedServiceUrl ?? "https://smba.trafficmanager.net/teams/";
    const tenantId = cachedTenantId ?? this.config.appTenantId;

    if (!tenantId) {
      throw new ValidationError(
        "teamssdk",
        "Cannot open DM: tenant ID not found. User must interact with the bot first."
      );
    }

    let conversationId = "";

    // biome-ignore lint/suspicious/noExplicitAny: BotBuilder types are incomplete
    await (this.botAdapter as any).createConversationAsync(
      this.config.appId,
      "msteams",
      serviceUrl,
      "",
      {
        isGroup: false,
        bot: { id: this.config.appId, name: this.userName },
        members: [{ id: userId }],
        tenantId,
        channelData: { tenant: { id: tenantId } },
      },
      async (turnContext: TurnContext) => {
        conversationId = turnContext?.activity?.conversation?.id ?? "";
      }
    );

    if (!conversationId) {
      throw new NetworkError(
        "teamssdk",
        "Failed to create 1:1 conversation - no ID returned"
      );
    }

    return this.encodeThreadId({ conversationId, serviceUrl });
  }

  async fetchMessages(
    threadId: string,
    options: FetchOptions = {}
  ): Promise<FetchResult<unknown>> {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "fetchMessages requires appTenantId for Microsoft Graph API access.",
        "fetchMessages"
      );
    }

    const { conversationId } = this.decodeThreadId(threadId);
    const limit = options.limit ?? 50;
    const cursor = options.cursor;
    const direction = options.direction ?? "backward";

    const messageIdMatch = conversationId.match(
      SEMICOLON_MESSAGEID_CAPTURE_PATTERN
    );
    const threadMessageId = messageIdMatch?.[1];
    const baseConversationId = conversationId.replace(
      MESSAGEID_STRIP_PATTERN,
      ""
    );

    // Check for cached channel context
    let channelContext: TeamsChannelContext | null = null;
    if (threadMessageId && this.chat) {
      const cachedContext = await this.chat
        .getState()
        .get<string>(`teamssdk:channelContext:${baseConversationId}`);
      if (cachedContext) {
        try {
          channelContext = JSON.parse(cachedContext) as TeamsChannelContext;
        } catch {
          // ignore invalid cache
        }
      }
    }

    try {
      if (channelContext && threadMessageId) {
        return this.fetchChannelThreadMessages(
          channelContext,
          threadMessageId,
          threadId,
          options
        );
      }

      let graphMessages: GraphChatMessage[];
      let hasMoreMessages = false;

      if (direction === "forward") {
        const allMessages: GraphChatMessage[] = [];
        let nextLink: string | undefined;
        const apiUrl = `/chats/${encodeURIComponent(baseConversationId)}/messages`;

        do {
          const request = nextLink
            ? this.graphClient.api(nextLink)
            : this.graphClient
                .api(apiUrl)
                .top(50)
                .orderby("createdDateTime desc");
          const response = await request.get() as { value?: GraphChatMessage[]; "@odata.nextLink"?: string };
          allMessages.push(...(response.value ?? []));
          nextLink = response["@odata.nextLink"];
        } while (nextLink);

        allMessages.reverse();

        let startIndex = 0;
        if (cursor) {
          startIndex = allMessages.findIndex(
            (msg) => msg.createdDateTime && msg.createdDateTime > cursor
          );
          if (startIndex === -1) startIndex = allMessages.length;
        }

        hasMoreMessages = startIndex + limit < allMessages.length;
        graphMessages = allMessages.slice(startIndex, startIndex + limit);
      } else {
        let request = this.graphClient
          .api(`/chats/${encodeURIComponent(baseConversationId)}/messages`)
          .top(limit)
          .orderby("createdDateTime desc");

        if (cursor) {
          request = request.filter(`createdDateTime lt ${cursor}`);
        }

        const response = await request.get() as { value?: GraphChatMessage[] };
        graphMessages = response.value ?? [];
        graphMessages.reverse();
        hasMoreMessages = graphMessages.length >= limit;
      }

      if (threadMessageId && !channelContext) {
        graphMessages = graphMessages.filter(
          (msg) => msg.id && msg.id >= threadMessageId
        );
      }

      const messages = this.graphMessagesToMessages(graphMessages, threadId);

      let nextCursor: string | undefined;
      if (hasMoreMessages && graphMessages.length > 0) {
        const refMsg =
          direction === "forward"
            ? graphMessages.at(-1)
            : graphMessages[0];
        if (refMsg?.createdDateTime) nextCursor = refMsg.createdDateTime;
      }

      return { messages, nextCursor };
    } catch (error) {
      this.logger.error("Teams Graph API: fetchMessages error", { error });
      if (error instanceof Error && error.message?.includes("403")) {
        throw new PermissionError(
          "teamssdk",
          "fetchMessages requires ChatMessage.Read.Chat or Chat.Read.All Graph permission"
        );
      }
      throw error;
    }
  }

  private async fetchChannelThreadMessages(
    channelContext: TeamsChannelContext,
    threadMessageId: string,
    threadId: string,
    options: FetchOptions
  ): Promise<FetchResult<unknown>> {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "fetchMessages requires graphClient",
        "fetchMessages"
      );
    }

    const limit = options.limit ?? 50;
    const cursor = options.cursor;

    const apiBase = `/teams/${encodeURIComponent(channelContext.teamId)}/channels/${encodeURIComponent(channelContext.channelId)}/messages/${encodeURIComponent(threadMessageId)}/replies`;

    let request = this.graphClient
      .api(apiBase)
      .top(limit)
      .orderby("createdDateTime desc");

    if (cursor) {
      request = request.filter(`createdDateTime lt ${cursor}`);
    }

    const response = await request.get() as { value?: GraphChatMessage[] };
    const graphMessages: GraphChatMessage[] = response.value ?? [];
    graphMessages.reverse();

    const messages = this.graphMessagesToMessages(graphMessages, threadId);
    let nextCursor: string | undefined;
    if (graphMessages.length >= limit) {
      const oldest = graphMessages[0];
      if (oldest?.createdDateTime) nextCursor = oldest.createdDateTime;
    }

    return { messages, nextCursor };
  }

  async fetchThread(threadId: string): Promise<ThreadInfo> {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);
    return {
      id: threadId,
      channelId: this.channelIdFromThreadId(threadId),
      isDM: this.isDM(threadId),
      metadata: { conversationId, serviceUrl },
    };
  }

  async fetchChannelInfo(channelId: string): Promise<ChannelInfo> {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "fetchChannelInfo requires Microsoft Graph API access.",
        "fetchChannelInfo"
      );
    }

    const parts = channelId.split(":");
    const teamId = parts[0];
    const channelPart = parts.slice(1).join(":");

    const response = await this.graphClient
      .api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}`)
      .get() as {
        id: string;
        displayName?: string;
        description?: string;
        membershipType?: string;
        createdDateTime?: string;
      };

    return {
      id: channelId,
      name: response.displayName ?? channelId,
      metadata: {
        description: response.description,
        isPrivate: response.membershipType === "private",
        createdDateTime: response.createdDateTime,
      },
    };
  }

  async fetchChannelMessages(
    channelId: string,
    options: FetchOptions = {}
  ): Promise<FetchResult<unknown>> {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "fetchChannelMessages requires Microsoft Graph API access.",
        "fetchChannelMessages"
      );
    }

    const parts = channelId.split(":");
    const teamId = parts[0];
    const channelPart = parts.slice(1).join(":");
    const limit = options.limit ?? 50;

    const response = await this.graphClient
      .api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}/messages`)
      .top(limit)
      .orderby("createdDateTime desc")
      .get() as { value?: GraphChatMessage[] };

    const graphMessages: GraphChatMessage[] = response.value ?? [];
    graphMessages.reverse();

    const threadId = `teamssdk:${channelId}:channel`;
    const messages = this.graphMessagesToMessages(graphMessages, threadId);

    return { messages };
  }

  async postChannelMessage(
    channelId: string,
    message: AdapterPostableMessage
  ): Promise<RawMessage<unknown>> {
    const parts = channelId.split(":");
    const teamId = parts[0];
    const channelPart = parts.slice(1).join(":");

    const card = extractCard(message);
    const threadId = `teamssdk:${channelId}:channel`;

    if (card) {
      const adaptiveCard = cardToAdaptiveCard(card);
      const payload = {
        body: { contentType: "html", content: "<attachment id=\"card1\"></attachment>" },
        attachments: [
          {
            id: "card1",
            contentType: "application/vnd.microsoft.card.adaptive",
            content: JSON.stringify(adaptiveCard),
          },
        ],
      };

      if (!this.graphClient) {
        throw new NotImplementedError(
          "postChannelMessage with cards requires Graph API access",
          "postChannelMessage"
        );
      }

      const response = await this.graphClient
        .api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}/messages`)
        .post(payload) as { id: string };

      return { id: response.id, threadId, raw: payload };
    }

    const text = convertEmojiPlaceholders(
      this.formatConverter.renderPostable(message),
      "teams"
    );

    if (!this.graphClient) {
      throw new NotImplementedError(
        "postChannelMessage requires Graph API access",
        "postChannelMessage"
      );
    }

    const payload = { body: { contentType: "html", content: text } };
    const response = await this.graphClient
      .api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}/messages`)
      .post(payload) as { id: string };

    return { id: response.id, threadId, raw: payload };
  }

  async listThreads(
    channelId: string,
    _options?: ListThreadsOptions
  ): Promise<ListThreadsResult<unknown>> {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "listThreads requires Microsoft Graph API access.",
        "listThreads"
      );
    }

    const parts = channelId.split(":");
    const teamId = parts[0];
    const channelPart = parts.slice(1).join(":");

    const response = await this.graphClient
      .api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}/messages`)
      .top(50)
      .orderby("createdDateTime desc")
      .get() as { value?: GraphChatMessage[] };

    const graphMessages: GraphChatMessage[] = response.value ?? [];

    const threads = graphMessages.map((msg) => {
      const rootThreadId = this.encodeThreadId({
        conversationId: `${channelId};messageid=${msg.id}`,
        serviceUrl: "",
      });
      const rawText = msg.body?.content?.replace(/<[^>]+>/g, "") ?? "";
      const rootMessage = new Message({
        id: msg.id,
        threadId: rootThreadId,
        text: rawText,
        formatted: this.formatConverter.toAst(rawText),
        raw: msg,
        author: {
          userId: msg.from?.user?.id ?? msg.from?.application?.id ?? "unknown",
          userName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? "unknown",
          fullName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? "unknown",
          isBot: !!msg.from?.application,
          isMe: msg.from?.application?.id === this.config.appId,
        },
        metadata: {
          dateSent: msg.createdDateTime ? new Date(msg.createdDateTime) : new Date(),
          edited: !!msg.lastModifiedDateTime,
        },
        attachments: [],
      });
      return {
        id: rootThreadId,
        rootMessage,
        lastReplyAt: msg.lastModifiedDateTime
          ? new Date(msg.lastModifiedDateTime)
          : undefined,
      };
    });

    return { threads };
  }

  parseMessage(raw: unknown): Message<unknown> {
    const activity = raw as Activity;
    const threadId = this.encodeThreadId({
      conversationId: activity.conversation?.id ?? "",
      serviceUrl: activity.serviceUrl ?? "",
    });
    return this.parseTeamsMessage(activity, threadId);
  }

  renderFormatted(content: FormattedContent): string {
    return this.formatConverter.fromAst(content);
  }

  encodeThreadId(data: TeamsThreadId): string {
    const encoded = Buffer.from(JSON.stringify(data)).toString("base64url");
    return `teamssdk:${encoded}`;
  }

  decodeThreadId(threadId: string): TeamsThreadId {
    const prefix = "teamssdk:";
    const encoded = threadId.startsWith(prefix)
      ? threadId.slice(prefix.length)
      : threadId;
    try {
      return JSON.parse(
        Buffer.from(encoded, "base64url").toString("utf8")
      ) as TeamsThreadId;
    } catch {
      return { conversationId: threadId, serviceUrl: "" };
    }
  }

  channelIdFromThreadId(threadId: string): string {
    try {
      const { conversationId } = this.decodeThreadId(threadId);
      const base = conversationId.replace(MESSAGEID_STRIP_PATTERN, "");
      return base;
    } catch {
      return threadId;
    }
  }

  isDM(threadId: string): boolean {
    try {
      const { conversationId } = this.decodeThreadId(threadId);
      // Personal/DM conversations don't have the @thread suffix
      return (
        !conversationId.includes("@thread") &&
        !conversationId.includes("@conference")
      );
    } catch {
      return false;
    }
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  private parseTeamsMessage(activity: Activity, threadId: string): Message<unknown> {
    const text = activity.text ?? "";
    const normalizedText = text.trim();
    const isMe = this.isMessageFromSelf(activity);

    return new Message({
      id: activity.id ?? "",
      threadId,
      text: this.formatConverter.extractPlainText(normalizedText),
      formatted: this.formatConverter.toAst(normalizedText),
      raw: activity,
      author: {
        userId: activity.from?.id ?? "unknown",
        userName: activity.from?.name ?? "unknown",
        fullName: activity.from?.name ?? "unknown",
        isBot: activity.from?.role === "bot",
        isMe,
      },
      metadata: {
        dateSent: activity.timestamp ? new Date(activity.timestamp) : new Date(),
        edited: false,
      },
      attachments: (activity.attachments ?? [])
        .filter(
          (att) =>
            att.contentType !== "application/vnd.microsoft.card.adaptive" &&
            !(att.contentType === "text/html" && !att.contentUrl)
        )
        .map((att) => this.createAttachment(att)),
    });
  }

  private createAttachment(att: {
    contentType?: string;
    contentUrl?: string;
    name?: string;
  }): Attachment {
    const url = att.contentUrl;

    let type: Attachment["type"] = "file";
    if (att.contentType?.startsWith("image/")) type = "image";
    else if (att.contentType?.startsWith("video/")) type = "video";
    else if (att.contentType?.startsWith("audio/")) type = "audio";

    return {
      type,
      url,
      name: att.name,
      mimeType: att.contentType,
      fetchData: url
        ? async () => {
            const response = await fetch(url);
            if (!response.ok) {
              throw new NetworkError(
                "teamssdk",
                `Failed to fetch file: ${response.status} ${response.statusText}`
              );
            }
            const arrayBuffer = await response.arrayBuffer();
            return Buffer.from(arrayBuffer);
          }
        : undefined,
    };
  }

  private graphMessagesToMessages(
    graphMessages: GraphChatMessage[],
    threadId: string
  ): Message<unknown>[] {
    return graphMessages.map((msg) => {
      const isFromBot =
        msg.from?.application?.id === this.config.appId ||
        msg.from?.user?.id === this.config.appId;
      const rawText = this.extractTextFromGraphMessage(msg);

      return new Message({
        id: msg.id,
        threadId,
        text: this.formatConverter.extractPlainText(rawText),
        formatted: this.formatConverter.toAst(rawText),
        raw: msg,
        author: {
          userId: msg.from?.user?.id ?? msg.from?.application?.id ?? "unknown",
          userName:
            msg.from?.user?.displayName ??
            msg.from?.application?.displayName ??
            "unknown",
          fullName:
            msg.from?.user?.displayName ??
            msg.from?.application?.displayName ??
            "unknown",
          isBot: !!msg.from?.application,
          isMe: isFromBot,
        },
        metadata: {
          dateSent: msg.createdDateTime
            ? new Date(msg.createdDateTime)
            : new Date(),
          edited: !!msg.lastModifiedDateTime,
        },
        attachments: this.extractAttachmentsFromGraphMessage(msg),
      });
    });
  }

  private extractTextFromGraphMessage(msg: GraphChatMessage): string {
    const body = msg.body;
    if (!body) return "";
    if (body.contentType === "html") {
      return body.content?.replace(/<[^>]+>/g, "") ?? "";
    }
    return body.content ?? "";
  }

  private extractAttachmentsFromGraphMessage(
    msg: GraphChatMessage
  ): Attachment[] {
    return (msg.attachments ?? [])
      .filter(
        (att) =>
          att.contentType !== "application/vnd.microsoft.card.adaptive"
      )
      .map((att) => ({
        type: "file" as const,
        url: att.contentUrl,
        name: att.name,
        mimeType: att.contentType,
      }));
  }

  private isMessageFromSelf(activity: Activity): boolean {
    return activity.from?.id === this.config.appId;
  }

  /** Map Teams API errors to typed adapter errors */
  private handleTeamsError(error: unknown, _operation: string): never {
    if (error instanceof Error) {
      const msg = error.message ?? "";
      if (msg.includes("429")) {
        const retryAfter = 10;
        throw new AdapterRateLimitError("teamssdk", retryAfter);
      }
      if (msg.includes("401")) {
        throw new AuthenticationError("teamssdk", msg);
      }
      if (msg.includes("403")) {
        throw new PermissionError("teamssdk", msg);
      }
    }
    throw error;
  }
}

/** Factory function for creating a TeamsSDKAdapter instance */
export function createTeamsSDKAdapter(
  config?: TeamsSDKAdapterConfig
): TeamsSDKAdapter {
  return new TeamsSDKAdapter(config);
}

export { cardToAdaptiveCard, cardToFallbackText } from "./cards.js";
export { TeamsSDKFormatConverter } from "./markdown.js";
