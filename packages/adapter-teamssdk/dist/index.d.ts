import type { Adapter, TeamsAdapter, TeamsAdapterConfig, TeamsActivity, TeamsApp, TeamsContext, ChatInstance, WebhookRequest, WebhookResponse, Message, SentMessage, Thread, ChannelInfo, Modal, FormattedContent } from './types.js';
import { adaptiveCardToAttachment } from './utils/format-converter.js';
/**
 * Microsoft Teams SDK Adapter implementing the vercel/chat Adapter interface.
 */
export declare class TeamsAdapterImpl implements Adapter, TeamsAdapter {
    readonly config: TeamsAdapterConfig;
    readonly app: TeamsApp;
    /** Exposed for handler modules to use; set after initialize(). */
    chatInstance?: ChatInstance;
    private readonly accessTokenCache;
    private readonly serviceUrlCache;
    private readonly logger;
    private readonly appImpl;
    constructor(config: TeamsAdapterConfig);
    initialize(chat: ChatInstance): Promise<void>;
    handleWebhook(request: WebhookRequest): Promise<WebhookResponse>;
    postMessage(threadId: string, message: Message): Promise<SentMessage>;
    editMessage(threadId: string, messageId: string, message: Message): Promise<void>;
    deleteMessage(threadId: string, messageId: string): Promise<void>;
    fetchMessages(threadId: string): Promise<Message[]>;
    fetchMessage(threadId: string, messageId: string): Promise<Message>;
    fetchChannelMessages(channelId: string): Promise<Message[]>;
    postChannelMessage(channelId: string, message: Message): Promise<SentMessage>;
    addReaction(threadId: string, messageId: string, emoji: string): Promise<void>;
    removeReaction(threadId: string, messageId: string, emoji: string): Promise<void>;
    fetchThread(threadId: string): Promise<Thread>;
    fetchChannelInfo(channelId: string): Promise<ChannelInfo>;
    listThreads(channelId: string): Promise<Thread[]>;
    openDM(userId: string): Promise<string>;
    openModal(_triggerId: string, modal: Modal): Promise<void>;
    startTyping(threadId: string): Promise<void>;
    encodeThreadId(context: TeamsContext): string;
    decodeThreadId(threadId: string): TeamsContext;
    channelIdFromThreadId(threadId: string): string;
    isDM(threadId: string): boolean;
    parseMessage(activity: TeamsActivity): Message;
    renderFormatted(content: FormattedContent): string;
    private ensureAccessToken;
    private buildCredentials;
    private botConnectorHeaders;
    private getServiceUrl;
    private createGraphClient;
    /**
     * Wraps an HTTP fetch call with retry logic and exponential back-off.
     * Respects rate-limit errors by honoring the retry-after value.
     */
    private callWithRetry;
    private registerInternalHandlers;
    private graphMessageToMessage;
}
/**
 * Creates and returns a new TeamsAdapterImpl instance.
 */
export declare function createTeamsAdapter(config: TeamsAdapterConfig): TeamsAdapterImpl;
export { adaptiveCardToAttachment };
export * from './types.js';
export { TeamsAdapterError, TeamsAdapterErrorCode } from './utils/error-handler.js';
export { encodeThreadId, decodeThreadId, channelIdFromThreadId, isDM as isDirectMessage } from './utils/thread-utils.js';
export { GraphClient } from './utils/graph-client.js';
//# sourceMappingURL=index.d.ts.map