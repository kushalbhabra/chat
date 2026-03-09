import { Activity, TurnContext } from 'botbuilder';
import { CardElement, BaseFormatConverter, AdapterPostableMessage, Root, Adapter, Logger, ChatInstance, WebhookOptions, RawMessage, EmojiValue, FetchOptions, FetchResult, ThreadInfo, ChannelInfo, ListThreadsOptions, ListThreadsResult, Message, FormattedContent } from 'chat';

/**
 * Teams Adaptive Card converter for cross-platform cards.
 *
 * Converts CardElement to Microsoft Adaptive Cards format.
 * @see https://adaptivecards.io/
 */

interface AdaptiveCard {
    $schema: string;
    actions?: AdaptiveCardAction[];
    body: AdaptiveCardElement[];
    type: "AdaptiveCard";
    version: string;
}
interface AdaptiveCardElement {
    type: string;
    [key: string]: unknown;
}
interface AdaptiveCardAction {
    data?: Record<string, unknown>;
    style?: string;
    title: string;
    type: string;
    url?: string;
}
/**
 * Convert a CardElement to a Teams Adaptive Card.
 */
declare function cardToAdaptiveCard(card: CardElement): AdaptiveCard;
/**
 * Generate fallback text from a card element.
 * Used when adaptive cards aren't supported.
 */
declare function cardToFallbackText(card: CardElement): string;

/**
 * Teams-specific format conversion using AST-based parsing.
 *
 * Teams supports a subset of HTML for formatting:
 * - Bold: <b> or <strong>
 * - Italic: <i> or <em>
 * - Strikethrough: <s> or <strike>
 * - Links: <a href="url">text</a>
 * - Code: <pre> and <code>
 *
 * Teams also accepts standard markdown in most cases.
 */

declare class TeamsSDKFormatConverter extends BaseFormatConverter {
    /**
     * Convert @mentions to Teams format in plain text.
     * @name → <at>name</at>
     */
    private convertMentionsToTeams;
    /**
     * Override renderPostable to convert @mentions in plain strings.
     */
    renderPostable(message: AdapterPostableMessage): string;
    /**
     * Render an AST to Teams format.
     * Teams accepts standard markdown, so we just stringify cleanly.
     */
    fromAst(ast: Root): string;
    /**
     * Parse Teams message into an AST.
     * Converts Teams HTML/mentions to standard markdown format.
     */
    toAst(teamsText: string): Root;
    private nodeToTeams;
    /**
     * Render an mdast table node as a GFM markdown table.
     * Teams renders markdown tables natively.
     */
    private tableToGfm;
}

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

/** Handler function type for Teams events */
type TeamsEventHandler = (activity: Activity, context: TurnContext) => Promise<void> | void;
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
declare class TeamsApp {
    private readonly handlers;
    /** Handle regular channel/group messages */
    $onMessage(handler: TeamsEventHandler): this;
    /** Handle @mention messages (the bot was explicitly mentioned) */
    $onMention(handler: TeamsEventHandler): this;
    /** Handle replies in a thread */
    $onThreadReplyAdded(handler: TeamsEventHandler): this;
    /** Handle 1:1 direct messages sent to the bot */
    $onDMReceived(handler: TeamsEventHandler): this;
    /** Handle emoji reactions added to a message */
    $onReactionAdded(handler: TeamsEventHandler): this;
    /** Handle emoji reactions removed from a message */
    $onReactionRemoved(handler: TeamsEventHandler): this;
    /** Handle Adaptive Card Action.Submit clicks */
    $onCardAction(handler: TeamsEventHandler): this;
    /** Handle invoke activities (task modules, adaptive card invokes) */
    $onInvoke(handler: TeamsEventHandler): this;
    /** Handle message context-menu actions */
    $onMessageAction(handler: TeamsEventHandler): this;
    /** Handle member added to team/conversation */
    $onMemberAdded(handler: TeamsEventHandler): this;
    /** Handle member removed from team/conversation */
    $onMemberRemoved(handler: TeamsEventHandler): this;
    /** Handle team rename event */
    $onTeamRenamed(handler: TeamsEventHandler): this;
    /** Handle channel created in a team */
    $onChannelCreated(handler: TeamsEventHandler): this;
    /** Handle channel renamed in a team */
    $onChannelRenamed(handler: TeamsEventHandler): this;
    /** Handle channel deleted from a team */
    $onChannelDeleted(handler: TeamsEventHandler): this;
    /** Handle app installation in a team/personal scope */
    $onAppInstalled(handler: TeamsEventHandler): this;
    /** Handle app uninstallation */
    $onAppUninstalled(handler: TeamsEventHandler): this;
    /** Handle any bot activity (fires for every activity, regardless of type) */
    $onBotActivity(handler: TeamsEventHandler): this;
    /**
     * Route an incoming activity to all matching registered handlers.
     * Always fires $onBotActivity handlers, then type-specific handlers.
     */
    processActivity(activity: Activity, context: TurnContext): Promise<void>;
    private dispatchMessageActivity;
    private dispatchReactionActivity;
    private dispatchInvokeActivity;
    private dispatchConversationUpdateActivity;
    private dispatchInstallationUpdateActivity;
    private addHandler;
    private runHandlers;
}
/** Certificate-based authentication config */
interface TeamsAuthCertificate {
    /** PEM-encoded certificate private key */
    certificatePrivateKey: string;
    /** Hex-encoded certificate thumbprint (optional when x5c is provided) */
    certificateThumbprint?: string;
    /** Public certificate for subject-name validation (optional) */
    x5c?: string;
}
/** Federated (workload identity) authentication config */
interface TeamsAuthFederated {
    /** Audience for the federated credential (defaults to api://AzureADTokenExchange) */
    clientAudience?: string;
    /** Client ID for the managed identity assigned to the bot */
    clientId: string;
}
interface TeamsSDKAdapterConfig {
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
interface TeamsThreadId {
    conversationId: string;
    replyToId?: string;
    serviceUrl: string;
}
declare class TeamsSDKAdapter implements Adapter<TeamsThreadId, unknown> {
    readonly name = "teamssdk";
    readonly userName: string;
    readonly botUserId?: string;
    /** The internal TeamsApp event router (teams.ts-style API) */
    readonly app: TeamsApp;
    private readonly botAdapter;
    private readonly graphClient;
    private chat;
    private readonly logger;
    private readonly formatConverter;
    private readonly config;
    constructor(config?: TeamsSDKAdapterConfig);
    /**
     * Wire up the TeamsApp event handlers to forward activities to the
     * Chat instance (processMessage / processAction / processReaction).
     */
    private wireAppHandlers;
    initialize(chat: ChatInstance): Promise<void>;
    handleWebhook(request: Request, options?: WebhookOptions): Promise<Response>;
    private handleTurn;
    private handleMessageActivity;
    private handleReactionActivity;
    private handleCardActionActivity;
    private handleInvokeActivity;
    postMessage(threadId: string, message: AdapterPostableMessage): Promise<RawMessage<unknown>>;
    private filesToAttachments;
    editMessage(threadId: string, messageId: string, message: AdapterPostableMessage): Promise<RawMessage<unknown>>;
    deleteMessage(threadId: string, messageId: string): Promise<void>;
    addReaction(_threadId: string, _messageId: string, _emoji: EmojiValue | string): Promise<void>;
    removeReaction(_threadId: string, _messageId: string, _emoji: EmojiValue | string): Promise<void>;
    startTyping(threadId: string, _status?: string): Promise<void>;
    openDM(userId: string): Promise<string>;
    fetchMessages(threadId: string, options?: FetchOptions): Promise<FetchResult<unknown>>;
    private fetchChannelThreadMessages;
    fetchThread(threadId: string): Promise<ThreadInfo>;
    fetchChannelInfo(channelId: string): Promise<ChannelInfo>;
    fetchChannelMessages(channelId: string, options?: FetchOptions): Promise<FetchResult<unknown>>;
    postChannelMessage(channelId: string, message: AdapterPostableMessage): Promise<RawMessage<unknown>>;
    listThreads(channelId: string, _options?: ListThreadsOptions): Promise<ListThreadsResult<unknown>>;
    parseMessage(raw: unknown): Message<unknown>;
    renderFormatted(content: FormattedContent): string;
    encodeThreadId(data: TeamsThreadId): string;
    decodeThreadId(threadId: string): TeamsThreadId;
    channelIdFromThreadId(threadId: string): string;
    isDM(threadId: string): boolean;
    private parseTeamsMessage;
    private createAttachment;
    private graphMessagesToMessages;
    private extractTextFromGraphMessage;
    private extractAttachmentsFromGraphMessage;
    private isMessageFromSelf;
    /** Map Teams API errors to typed adapter errors */
    private handleTeamsError;
}
/** Factory function for creating a TeamsSDKAdapter instance */
declare function createTeamsSDKAdapter(config?: TeamsSDKAdapterConfig): TeamsSDKAdapter;

export { TeamsApp, type TeamsAuthCertificate, type TeamsAuthFederated, type TeamsEventHandler, TeamsSDKAdapter, type TeamsSDKAdapterConfig, TeamsSDKFormatConverter, type TeamsThreadId, cardToAdaptiveCard, cardToFallbackText, createTeamsSDKAdapter };
