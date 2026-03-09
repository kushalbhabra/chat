export interface Adapter {
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
    encodeThreadId(context: TeamsContext): string;
    decodeThreadId(threadId: string): TeamsContext;
    channelIdFromThreadId(threadId: string): string;
    isDM(threadId: string): boolean;
    openDM(userId: string): Promise<string>;
    openModal(triggerId: string, modal: Modal): Promise<void>;
    startTyping(threadId: string): Promise<void>;
    fetchChannelInfo(channelId: string): Promise<ChannelInfo>;
    listThreads(channelId: string): Promise<Thread[]>;
    parseMessage(activity: TeamsActivity): Message;
    renderFormatted(content: FormattedContent): string;
}
export interface TeamsContext {
    serviceUrl: string;
    tenantId: string;
    conversationId: string;
    channelId?: string;
    teamId?: string;
    isDM?: boolean;
}
export interface TeamsAdapterConfig {
    appId: string;
    appPassword?: string;
    appCertificate?: {
        thumbprint: string;
        privateKey: string;
    };
    tenantId?: string;
    allowedTenants?: string[];
    enableLogging?: boolean;
    maxRetries?: number;
    retryDelayMs?: number;
    graphApiBaseUrl?: string;
    botFrameworkApiUrl?: string;
}
export interface TeamsActivity {
    type: string;
    id?: string;
    timestamp?: string;
    localTimestamp?: string;
    serviceUrl: string;
    channelId: string;
    from: {
        id: string;
        name?: string;
        aadObjectId?: string;
    };
    conversation: {
        id: string;
        isGroup?: boolean;
        conversationType?: string;
        tenantId?: string;
        name?: string;
    };
    recipient: {
        id: string;
        name?: string;
    };
    text?: string;
    textFormat?: string;
    attachments?: TeamsAttachment[];
    entities?: TeamsEntity[];
    channelData?: Record<string, unknown>;
    value?: unknown;
    replyToId?: string;
    membersAdded?: {
        id: string;
        name?: string;
    }[];
    membersRemoved?: {
        id: string;
        name?: string;
    }[];
    reactionsAdded?: {
        type: string;
    }[];
    reactionsRemoved?: {
        type: string;
    }[];
    name?: string;
}
export interface TeamsAttachment {
    contentType: string;
    content?: unknown;
    contentUrl?: string;
    name?: string;
    thumbnailUrl?: string;
}
export interface TeamsEntity {
    type: string;
    mentioned?: {
        id: string;
        name?: string;
    };
    text?: string;
    [key: string]: unknown;
}
export interface Message {
    id?: string;
    text?: string;
    formattedContent?: FormattedContent;
    userId?: string;
    userName?: string;
    timestamp?: string;
    attachments?: Attachment[];
    reactions?: Reaction[];
    threadId?: string;
    channelId?: string;
    replyToId?: string;
    metadata?: Record<string, unknown>;
}
export interface SentMessage {
    id: string;
    timestamp?: string;
}
export interface FormattedContent {
    type: 'text' | 'markdown' | 'html' | 'adaptive_card' | 'blocks';
    content: string | Record<string, unknown>;
    blocks?: Block[];
}
export interface Block {
    type: string;
    text?: string;
    elements?: unknown[];
    [key: string]: unknown;
}
export interface Attachment {
    type: 'image' | 'file' | 'card' | 'embed';
    url?: string;
    name?: string;
    content?: unknown;
    contentType?: string;
}
export interface Reaction {
    emoji: string;
    count: number;
    users?: string[];
}
export interface Thread {
    id: string;
    channelId: string;
    createdAt?: string;
    lastActivity?: string;
    messageCount?: number;
    topic?: string;
    isDM?: boolean;
}
export interface ChannelInfo {
    id: string;
    name?: string;
    teamId?: string;
    description?: string;
    isPrivate?: boolean;
    createdAt?: string;
    memberCount?: number;
}
export interface Modal {
    title: string;
    submitLabel?: string;
    cancelLabel?: string;
    adaptiveCard?: Record<string, unknown>;
    fields?: ModalField[];
}
export interface ModalField {
    id: string;
    type: 'text' | 'select' | 'checkbox' | 'date';
    label: string;
    placeholder?: string;
    required?: boolean;
    options?: {
        label: string;
        value: string;
    }[];
}
export interface WebhookRequest {
    headers: Record<string, string>;
    body: unknown;
    method?: string;
    url?: string;
}
export interface WebhookResponse {
    status: number;
    body?: unknown;
    headers?: Record<string, string>;
}
export interface ChatInstance {
    emit(event: string, data: unknown): void;
    on(event: string, handler: (data: unknown) => void): void;
}
export type TeamsEventHandler = (activity: TeamsActivity, adapter: TeamsAdapter) => Promise<void> | void;
export interface TeamsApp {
    $onMessage: (handler: TeamsEventHandler) => void;
    $onMention: (handler: TeamsEventHandler) => void;
    $onThreadReplyAdded: (handler: TeamsEventHandler) => void;
    $onDMReceived: (handler: TeamsEventHandler) => void;
    $onReactionAdded: (handler: TeamsEventHandler) => void;
    $onReactionRemoved: (handler: TeamsEventHandler) => void;
    $onCardAction: (handler: TeamsEventHandler) => void;
    $onInvoke: (handler: TeamsEventHandler) => void;
    $onMessageAction: (handler: TeamsEventHandler) => void;
    $onMemberAdded: (handler: TeamsEventHandler) => void;
    $onMemberRemoved: (handler: TeamsEventHandler) => void;
    $onTeamRenamed: (handler: TeamsEventHandler) => void;
    $onChannelCreated: (handler: TeamsEventHandler) => void;
    $onChannelRenamed: (handler: TeamsEventHandler) => void;
    $onChannelDeleted: (handler: TeamsEventHandler) => void;
    $onAppInstalled: (handler: TeamsEventHandler) => void;
    $onAppUninstalled: (handler: TeamsEventHandler) => void;
    $onBotActivity: (handler: TeamsEventHandler) => void;
    processActivity(activity: TeamsActivity): Promise<void>;
}
export interface TokenCredentials {
    type: 'password' | 'certificate' | 'federated';
    appId: string;
    password?: string;
    thumbprint?: string;
    privateKey?: string;
    federatedToken?: string;
}
export interface AccessToken {
    token: string;
    expiresAt: number;
}
export interface RateLimitError extends Error {
    retryAfter: number;
}
export interface TeamsAdapter extends Adapter {
    config: TeamsAdapterConfig;
    app: TeamsApp;
}
//# sourceMappingURL=types.d.ts.map