import type {
  Adapter,
  TeamsAdapter,
  TeamsAdapterConfig,
  TeamsActivity,
  TeamsApp,
  TeamsContext,
  TeamsEventHandler,
  ChatInstance,
  WebhookRequest,
  WebhookResponse,
  Message,
  SentMessage,
  Thread,
  ChannelInfo,
  Modal,
  FormattedContent,
  AccessToken,
  TokenCredentials,
} from './types.js';
import {
  encodeThreadId,
  decodeThreadId,
  channelIdFromThreadId,
  isDM as isDMThread,
} from './utils/thread-utils.js';
import { getAccessToken, validateToken } from './utils/auth-validator.js';
import {
  activityToMessage,
  messageToActivity,
  renderMarkdown,
  adaptiveCardToAttachment,
} from './utils/format-converter.js';
import {
  TeamsAdapterError,
  TeamsAdapterErrorCode,
  handleApiError,
  isRateLimitError,
  handleHttpError,
} from './utils/error-handler.js';
import { GraphClient } from './utils/graph-client.js';
import { createMessageHandlers } from './handlers/message.js';
import { createCardActionHandlers } from './handlers/card-action.js';
import { createChannelEventHandlers } from './handlers/channel-event.js';
import { createLifecycleHandlers } from './handlers/lifecycle.js';

// ---------------------------------------------------------------------------
// Logger
// ---------------------------------------------------------------------------

interface Logger {
  log(message: string, ...args: unknown[]): void;
  warn(message: string, ...args: unknown[]): void;
  error(message: string, ...args: unknown[]): void;
}

function createLogger(enabled: boolean): Logger {
  if (!enabled) {
    return {
      log: () => undefined,
      warn: () => undefined,
      error: () => undefined,
    };
  }
  return {
    log: (msg, ...args) => console.log(`[TeamsAdapter] ${msg}`, ...args),
    warn: (msg, ...args) => console.warn(`[TeamsAdapter] WARN ${msg}`, ...args),
    error: (msg, ...args) => console.error(`[TeamsAdapter] ERROR ${msg}`, ...args),
  };
}

// ---------------------------------------------------------------------------
// Internal TeamsApp implementation
// ---------------------------------------------------------------------------

class TeamsAppImpl implements TeamsApp {
  private readonly handlers = new Map<string, TeamsEventHandler[]>();

  private register(event: string, handler: TeamsEventHandler): void {
    const existing = this.handlers.get(event) ?? [];
    existing.push(handler);
    this.handlers.set(event, existing);
  }

  private async dispatch(
    event: string,
    activity: TeamsActivity,
    adapter: TeamsAdapter
  ): Promise<void> {
    const eventHandlers = this.handlers.get(event) ?? [];
    for (const handler of eventHandlers) {
      await handler(activity, adapter);
    }
  }

  $onMessage(handler: TeamsEventHandler): void { this.register('message', handler); }
  $onMention(handler: TeamsEventHandler): void { this.register('mention', handler); }
  $onThreadReplyAdded(handler: TeamsEventHandler): void { this.register('threadReplyAdded', handler); }
  $onDMReceived(handler: TeamsEventHandler): void { this.register('dmReceived', handler); }
  $onReactionAdded(handler: TeamsEventHandler): void { this.register('reactionAdded', handler); }
  $onReactionRemoved(handler: TeamsEventHandler): void { this.register('reactionRemoved', handler); }
  $onCardAction(handler: TeamsEventHandler): void { this.register('cardAction', handler); }
  $onInvoke(handler: TeamsEventHandler): void { this.register('invoke', handler); }
  $onMessageAction(handler: TeamsEventHandler): void { this.register('messageAction', handler); }
  $onMemberAdded(handler: TeamsEventHandler): void { this.register('memberAdded', handler); }
  $onMemberRemoved(handler: TeamsEventHandler): void { this.register('memberRemoved', handler); }
  $onTeamRenamed(handler: TeamsEventHandler): void { this.register('teamRenamed', handler); }
  $onChannelCreated(handler: TeamsEventHandler): void { this.register('channelCreated', handler); }
  $onChannelRenamed(handler: TeamsEventHandler): void { this.register('channelRenamed', handler); }
  $onChannelDeleted(handler: TeamsEventHandler): void { this.register('channelDeleted', handler); }
  $onAppInstalled(handler: TeamsEventHandler): void { this.register('appInstalled', handler); }
  $onAppUninstalled(handler: TeamsEventHandler): void { this.register('appUninstalled', handler); }
  $onBotActivity(handler: TeamsEventHandler): void { this.register('botActivity', handler); }

  async processActivity(activity: TeamsActivity, adapter?: TeamsAdapter): Promise<void> {
    if (!adapter) return;

    // Always emit the catch-all event
    await this.dispatch('botActivity', activity, adapter);

    switch (activity.type) {
      case 'message':
        await this.routeMessageActivity(activity, adapter);
        break;

      case 'messageReaction':
        if ((activity.reactionsAdded?.length ?? 0) > 0) {
          await this.dispatch('reactionAdded', activity, adapter);
        }
        if ((activity.reactionsRemoved?.length ?? 0) > 0) {
          await this.dispatch('reactionRemoved', activity, adapter);
        }
        break;

      case 'invoke':
        if (activity.name === 'composeExtension/submitAction' || activity.name === 'task/submit') {
          await this.dispatch('cardAction', activity, adapter);
        } else if (activity.name?.startsWith('message/submitAction')) {
          await this.dispatch('messageAction', activity, adapter);
        } else {
          await this.dispatch('invoke', activity, adapter);
        }
        break;

      case 'conversationUpdate':
        await this.routeConversationUpdate(activity, adapter);
        break;

      case 'installationUpdate':
        if (activity.name === 'add') {
          await this.dispatch('appInstalled', activity, adapter);
        } else if (activity.name === 'remove') {
          await this.dispatch('appUninstalled', activity, adapter);
        }
        break;

      default:
        // Unknown activity type – already emitted via botActivity above
        break;
    }
  }

  private async routeMessageActivity(
    activity: TeamsActivity,
    adapter: TeamsAdapter
  ): Promise<void> {
    const isMention = (activity.entities ?? []).some(
      (e) => e.type === 'mention' && e.mentioned?.id === activity.recipient.id
    );
    const isDM =
      activity.conversation.conversationType === 'personal' ||
      !activity.conversation.isGroup;
    const isReply = Boolean(activity.replyToId);

    if (isDM) {
      await this.dispatch('dmReceived', activity, adapter);
    } else if (isMention) {
      await this.dispatch('mention', activity, adapter);
    } else if (isReply) {
      await this.dispatch('threadReplyAdded', activity, adapter);
    } else {
      await this.dispatch('message', activity, adapter);
    }
  }

  private async routeConversationUpdate(
    activity: TeamsActivity,
    adapter: TeamsAdapter
  ): Promise<void> {
    const channelData = activity.channelData as Record<string, unknown> | undefined;
    const eventType = channelData?.['eventType'] as string | undefined;

    if ((activity.membersAdded?.length ?? 0) > 0) {
      await this.dispatch('memberAdded', activity, adapter);
    }
    if ((activity.membersRemoved?.length ?? 0) > 0) {
      await this.dispatch('memberRemoved', activity, adapter);
    }

    switch (eventType) {
      case 'teamRenamed':
        await this.dispatch('teamRenamed', activity, adapter);
        break;
      case 'channelCreated':
        await this.dispatch('channelCreated', activity, adapter);
        break;
      case 'channelRenamed':
        await this.dispatch('channelRenamed', activity, adapter);
        break;
      case 'channelDeleted':
        await this.dispatch('channelDeleted', activity, adapter);
        break;
      default:
        break;
    }
  }
}

// ---------------------------------------------------------------------------
// Main adapter class
// ---------------------------------------------------------------------------

/**
 * Microsoft Teams SDK Adapter implementing the vercel/chat Adapter interface.
 */
export class TeamsAdapterImpl implements Adapter, TeamsAdapter {
  readonly config: TeamsAdapterConfig;
  readonly app: TeamsApp;

  /** Exposed for handler modules to use; set after initialize(). */
  chatInstance?: ChatInstance;

  private readonly accessTokenCache = new Map<string, AccessToken>();
  private readonly serviceUrlCache = new Map<string, string>();
  private readonly logger: Logger;
  private readonly appImpl: TeamsAppImpl;

  constructor(config: TeamsAdapterConfig) {
    validateConfig(config);
    this.config = { ...config };
    this.logger = createLogger(config.enableLogging ?? false);
    this.appImpl = new TeamsAppImpl();
    this.app = this.appImpl;

    // Register internal handlers that bridge app events → ChatInstance events
    this.registerInternalHandlers();
  }

  // ---------------------------------------------------------------------------
  // Adapter interface – lifecycle
  // ---------------------------------------------------------------------------

  async initialize(chat: ChatInstance): Promise<void> {
    this.chatInstance = chat;
    this.logger.log('Adapter initialized');
  }

  // ---------------------------------------------------------------------------
  // Adapter interface – webhook
  // ---------------------------------------------------------------------------

  async handleWebhook(request: WebhookRequest): Promise<WebhookResponse> {
    try {
      // Validate the Authorization header
      const authHeader =
        request.headers['authorization'] ?? request.headers['Authorization'] ?? '';
      const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';

      if (!token) {
        this.logger.warn('Missing or malformed Authorization header');
        return { status: 401, body: { error: 'Unauthorized' } };
      }

      const isValid = await validateToken(token, this.config.appId);
      if (!isValid) {
        this.logger.warn('JWT validation failed');
        return { status: 401, body: { error: 'Unauthorized' } };
      }

      const activity = request.body as TeamsActivity;
      if (!activity || !activity.type) {
        return { status: 400, body: { error: 'Invalid activity payload' } };
      }

      // Cache the service URL for later outbound calls
      if (activity.serviceUrl) {
        this.serviceUrlCache.set(activity.conversation.id, activity.serviceUrl);
      }

      this.logger.log(`Processing activity type: ${activity.type}`);
      await this.appImpl.processActivity(activity, this);

      return { status: 200, body: {} };
    } catch (err) {
      this.logger.error('handleWebhook error', err);
      const adapted = handleApiError(err);
      return { status: adapted.statusCode, body: { error: adapted.message } };
    }
  }

  // ---------------------------------------------------------------------------
  // Adapter interface – messaging
  // ---------------------------------------------------------------------------

  async postMessage(threadId: string, message: Message): Promise<SentMessage> {
    const context = decodeThreadId(threadId);
    const serviceUrl = this.getServiceUrl(context);
    const activity = messageToActivity(message);
    activity.type = 'message';

    const url = `${serviceUrl}/v3/conversations/${encodeURIComponent(context.conversationId)}/activities`;
    const token = await this.ensureAccessToken();

    const response = await this.callWithRetry<{ id: string; timestamp?: string }>(() =>
      fetch(url, {
        method: 'POST',
        headers: this.botConnectorHeaders(token),
        body: JSON.stringify(activity),
      })
    );

    this.logger.log(`Posted message, id=${response.id}`);
    return { id: response.id, timestamp: response.timestamp };
  }

  async editMessage(
    threadId: string,
    messageId: string,
    message: Message
  ): Promise<void> {
    const context = decodeThreadId(threadId);
    const serviceUrl = this.getServiceUrl(context);
    const activity = messageToActivity(message);
    activity.type = 'message';
    activity.id = messageId;

    const url = `${serviceUrl}/v3/conversations/${encodeURIComponent(context.conversationId)}/activities/${encodeURIComponent(messageId)}`;
    const token = await this.ensureAccessToken();

    await this.callWithRetry<unknown>(() =>
      fetch(url, {
        method: 'PUT',
        headers: this.botConnectorHeaders(token),
        body: JSON.stringify(activity),
      })
    );
    this.logger.log(`Edited message id=${messageId}`);
  }

  async deleteMessage(threadId: string, messageId: string): Promise<void> {
    const context = decodeThreadId(threadId);
    const serviceUrl = this.getServiceUrl(context);
    const url = `${serviceUrl}/v3/conversations/${encodeURIComponent(context.conversationId)}/activities/${encodeURIComponent(messageId)}`;
    const token = await this.ensureAccessToken();

    await this.callWithRetry<unknown>(() =>
      fetch(url, {
        method: 'DELETE',
        headers: this.botConnectorHeaders(token),
      })
    );
    this.logger.log(`Deleted message id=${messageId}`);
  }

  async fetchMessages(threadId: string): Promise<Message[]> {
    const context = decodeThreadId(threadId);
    if (!context.teamId || !context.channelId) {
      throw new TeamsAdapterError(
        'fetchMessages requires teamId and channelId in the thread context',
        TeamsAdapterErrorCode.VALIDATION_ERROR,
        400
      );
    }
    const graphClient = await this.createGraphClient();
    const raw = await graphClient.getChannelMessages(context.teamId, context.channelId);
    return raw.map((item) => this.graphMessageToMessage(item as GraphMessage));
  }

  async fetchMessage(threadId: string, messageId: string): Promise<Message> {
    const context = decodeThreadId(threadId);
    if (!context.teamId || !context.channelId) {
      throw new TeamsAdapterError(
        'fetchMessage requires teamId and channelId in the thread context',
        TeamsAdapterErrorCode.VALIDATION_ERROR,
        400
      );
    }
    const graphClient = await this.createGraphClient();
    const raw = await graphClient.getChannelMessage(
      context.teamId,
      context.channelId,
      messageId
    );
    return this.graphMessageToMessage(raw as GraphMessage);
  }

  async fetchChannelMessages(channelId: string): Promise<Message[]> {
    // channelId here is expected to be a teamId/channelId composite "teamId:channelId"
    const [teamId, chanId] = parseCompositeChannelId(channelId);
    const graphClient = await this.createGraphClient();
    const raw = await graphClient.getChannelMessages(teamId, chanId);
    return raw.map((item) => this.graphMessageToMessage(item as GraphMessage));
  }

  async postChannelMessage(channelId: string, message: Message): Promise<SentMessage> {
    const [teamId, chanId] = parseCompositeChannelId(channelId);
    const graphClient = await this.createGraphClient();
    const body = buildGraphMessageBody(message);
    const result = (await graphClient.sendChatMessage(
      `${teamId}/channels/${chanId}`,
      body
    )) as { id?: string; createdDateTime?: string };
    return { id: result.id ?? '', timestamp: result.createdDateTime };
  }

  // ---------------------------------------------------------------------------
  // Adapter interface – reactions
  // ---------------------------------------------------------------------------

  async addReaction(
    threadId: string,
    messageId: string,
    emoji: string
  ): Promise<void> {
    const context = decodeThreadId(threadId);
    if (!context.teamId || !context.channelId) {
      throw new TeamsAdapterError(
        'addReaction requires teamId and channelId in the thread context',
        TeamsAdapterErrorCode.VALIDATION_ERROR,
        400
      );
    }
    const graphClient = await this.createGraphClient();
    await graphClient.setReaction(context.teamId, context.channelId, messageId, emoji);
    this.logger.log(`Added reaction ${emoji} to message ${messageId}`);
  }

  async removeReaction(
    threadId: string,
    messageId: string,
    emoji: string
  ): Promise<void> {
    const context = decodeThreadId(threadId);
    if (!context.teamId || !context.channelId) {
      throw new TeamsAdapterError(
        'removeReaction requires teamId and channelId in the thread context',
        TeamsAdapterErrorCode.VALIDATION_ERROR,
        400
      );
    }
    const graphClient = await this.createGraphClient();
    await graphClient.unsetReaction(context.teamId, context.channelId, messageId, emoji);
    this.logger.log(`Removed reaction ${emoji} from message ${messageId}`);
  }

  // ---------------------------------------------------------------------------
  // Adapter interface – threads / channels
  // ---------------------------------------------------------------------------

  async fetchThread(threadId: string): Promise<Thread> {
    const context = decodeThreadId(threadId);
    return {
      id: threadId,
      channelId: context.channelId ?? context.conversationId,
      isDM: context.isDM ?? false,
      topic: undefined,
    };
  }

  async fetchChannelInfo(channelId: string): Promise<ChannelInfo> {
    const [teamId, chanId] = parseCompositeChannelId(channelId);
    const graphClient = await this.createGraphClient();
    const raw = (await graphClient.getChannel(teamId, chanId)) as GraphChannel;
    return {
      id: raw.id ?? chanId,
      name: raw.displayName,
      teamId,
      description: raw.description,
      isPrivate: raw.membershipType === 'private',
      createdAt: raw.createdDateTime,
    };
  }

  async listThreads(channelId: string): Promise<Thread[]> {
    const [teamId, chanId] = parseCompositeChannelId(channelId);
    const graphClient = await this.createGraphClient();
    const messages = await graphClient.getChannelMessages(teamId, chanId);
    // Top-level messages that have replies are "threads"
    return messages
      .filter((m) => {
        const msg = m as GraphMessage;
        return msg.replyToId === null || msg.replyToId === undefined;
      })
      .map((m) => {
        const msg = m as GraphMessage;
        return {
          id: encodeThreadId({
            serviceUrl: '',
            tenantId: '',
            conversationId: msg.id ?? '',
            teamId,
            channelId: chanId,
          }),
          channelId: chanId,
          createdAt: msg.createdDateTime,
          lastActivity: msg.lastModifiedDateTime,
          topic: stripHtml(msg.body?.content ?? '').slice(0, 100),
        };
      });
  }

  // ---------------------------------------------------------------------------
  // Adapter interface – DM / modal / typing
  // ---------------------------------------------------------------------------

  async openDM(userId: string): Promise<string> {
    const graphClient = await this.createGraphClient();
    const chat = (await graphClient.createOneOnOneChat(userId, this.config.appId)) as {
      id?: string;
    };
    if (!chat.id) {
      throw new TeamsAdapterError('Failed to create DM chat', TeamsAdapterErrorCode.API_ERROR, 500);
    }
    // Return the chat ID as a minimal encoded thread ID
    const context: TeamsContext = {
      serviceUrl: this.config.botFrameworkApiUrl ?? 'https://smba.trafficmanager.net',
      tenantId: this.config.tenantId ?? '',
      conversationId: chat.id,
      isDM: true,
    };
    return encodeThreadId(context);
  }

  async openModal(_triggerId: string, modal: Modal): Promise<void> {
    // In Teams, modals are opened via a task/fetch invoke response.
    // The triggerId is the invoke activity ID; callers must reply via the
    // Bot Connector API. Here we log the intent; real usage requires the
    // invokeReply helper or a middleware integration.
    this.logger.log('openModal called – send a task/fetch response to the invoke activity', {
      title: modal.title,
    });
  }

  async startTyping(threadId: string): Promise<void> {
    const context = decodeThreadId(threadId);
    const serviceUrl = this.getServiceUrl(context);
    const url = `${serviceUrl}/v3/conversations/${encodeURIComponent(context.conversationId)}/activities`;
    const token = await this.ensureAccessToken();

    await this.callWithRetry<unknown>(() =>
      fetch(url, {
        method: 'POST',
        headers: this.botConnectorHeaders(token),
        body: JSON.stringify({ type: 'typing' }),
      })
    );
    this.logger.log(`Sent typing indicator to ${context.conversationId}`);
  }

  // ---------------------------------------------------------------------------
  // Adapter interface – ID helpers
  // ---------------------------------------------------------------------------

  encodeThreadId(context: TeamsContext): string {
    return encodeThreadId(context);
  }

  decodeThreadId(threadId: string): TeamsContext {
    return decodeThreadId(threadId);
  }

  channelIdFromThreadId(threadId: string): string {
    return channelIdFromThreadId(threadId);
  }

  isDM(threadId: string): boolean {
    return isDMThread(threadId);
  }

  // ---------------------------------------------------------------------------
  // Adapter interface – message parsing / rendering
  // ---------------------------------------------------------------------------

  parseMessage(activity: TeamsActivity): Message {
    return activityToMessage(activity);
  }

  renderFormatted(content: FormattedContent): string {
    return renderMarkdown(content);
  }

  // ---------------------------------------------------------------------------
  // Private: token management
  // ---------------------------------------------------------------------------

  private async ensureAccessToken(): Promise<string> {
    const cacheKey = this.config.appId;
    const cached = this.accessTokenCache.get(cacheKey);
    if (cached && cached.expiresAt > Date.now()) {
      return cached.token;
    }

    const credentials = this.buildCredentials();
    const token = await getAccessToken(credentials);
    this.accessTokenCache.set(cacheKey, token);
    return token.token;
  }

  private buildCredentials(): TokenCredentials {
    if (this.config.appCertificate) {
      return {
        type: 'certificate',
        appId: this.config.appId,
        thumbprint: this.config.appCertificate.thumbprint,
        privateKey: this.config.appCertificate.privateKey,
      };
    }
    if (this.config.appPassword) {
      return {
        type: 'password',
        appId: this.config.appId,
        password: this.config.appPassword,
      };
    }
    throw new TeamsAdapterError(
      'No credentials configured. Provide appPassword or appCertificate.',
      TeamsAdapterErrorCode.UNAUTHORIZED,
      401
    );
  }

  // ---------------------------------------------------------------------------
  // Private: HTTP helpers
  // ---------------------------------------------------------------------------

  private botConnectorHeaders(token: string): Record<string, string> {
    return {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
    };
  }

  private getServiceUrl(context: TeamsContext): string {
    const cached = this.serviceUrlCache.get(context.conversationId);
    if (cached) return cached;
    const url = context.serviceUrl ||
      this.config.botFrameworkApiUrl ||
      'https://smba.trafficmanager.net/apis';
    return url.replace(/\/$/, '');
  }

  private async createGraphClient(): Promise<GraphClient> {
    const token = await this.ensureAccessToken();
    return new GraphClient(token, this.config.graphApiBaseUrl);
  }

  /**
   * Wraps an HTTP fetch call with retry logic and exponential back-off.
   * Respects rate-limit errors by honoring the retry-after value.
   */
  private async callWithRetry<T>(
    fn: () => Promise<Response>,
    retries = this.config.maxRetries ?? 3,
    delayMs = this.config.retryDelayMs ?? 500
  ): Promise<T> {
    let lastError: unknown;
    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        const res = await fn();
        if (!res.ok) {
          const err = await handleHttpError(res);
          if (isRateLimitError(err)) {
            const retryAfter =
              Number(res.headers.get('retry-after') ?? '1') * 1000;
            this.logger.warn(`Rate limited, retrying after ${retryAfter}ms`);
            await sleep(retryAfter);
            lastError = err;
            continue;
          }
          throw err;
        }
        // 204 No Content
        if (res.status === 204) return {} as T;
        return (await res.json()) as T;
      } catch (err) {
        if (err instanceof TeamsAdapterError && !isTransient(err)) {
          throw err;
        }
        lastError = err;
        if (attempt < retries) {
          const backoff = delayMs * Math.pow(2, attempt);
          this.logger.warn(`Retrying (attempt ${attempt + 1}/${retries}) after ${backoff}ms`);
          await sleep(backoff);
        }
      }
    }
    throw handleApiError(lastError);
  }

  // ---------------------------------------------------------------------------
  // Private: internal handler registration
  // ---------------------------------------------------------------------------

  private registerInternalHandlers(): void {
    const msgHandlers = createMessageHandlers(this);
    const cardHandlers = createCardActionHandlers(this);
    const channelHandlers = createChannelEventHandlers(this);
    const lifecycleHandlers = createLifecycleHandlers(this);

    this.appImpl.$onMessage((a) => msgHandlers.onMessage(a));
    this.appImpl.$onMention((a) => msgHandlers.onMention(a));
    this.appImpl.$onThreadReplyAdded((a) => msgHandlers.onThreadReplyAdded(a));
    this.appImpl.$onDMReceived((a) => msgHandlers.onDMReceived(a));
    this.appImpl.$onReactionAdded((a) => msgHandlers.onReactionAdded(a));
    this.appImpl.$onReactionRemoved((a) => msgHandlers.onReactionRemoved(a));
    this.appImpl.$onCardAction((a) => cardHandlers.onCardAction(a));
    this.appImpl.$onInvoke((a) => cardHandlers.onInvoke(a));
    this.appImpl.$onMessageAction((a) => cardHandlers.onMessageAction(a));
    this.appImpl.$onMemberAdded((a) => channelHandlers.onMemberAdded(a));
    this.appImpl.$onMemberRemoved((a) => channelHandlers.onMemberRemoved(a));
    this.appImpl.$onTeamRenamed((a) => channelHandlers.onTeamRenamed(a));
    this.appImpl.$onChannelCreated((a) => channelHandlers.onChannelCreated(a));
    this.appImpl.$onChannelRenamed((a) => channelHandlers.onChannelRenamed(a));
    this.appImpl.$onChannelDeleted((a) => channelHandlers.onChannelDeleted(a));
    this.appImpl.$onAppInstalled((a) => lifecycleHandlers.onAppInstalled(a));
    this.appImpl.$onAppUninstalled((a) => lifecycleHandlers.onAppUninstalled(a));
    this.appImpl.$onBotActivity((a) => lifecycleHandlers.onBotActivity(a));
  }

  // ---------------------------------------------------------------------------
  // Private: Graph message normalisation
  // ---------------------------------------------------------------------------

  private graphMessageToMessage(raw: GraphMessage): Message {
    return {
      id: raw.id,
      text: stripHtml(raw.body?.content ?? ''),
      userId: raw.from?.user?.id,
      userName: raw.from?.user?.displayName,
      timestamp: raw.createdDateTime,
      threadId: raw.id,
      replyToId: raw.replyToId ?? undefined,
      metadata: {
        etag: raw.etag,
        messageType: raw.messageType,
      },
    };
  }
}

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

/**
 * Creates and returns a new TeamsAdapterImpl instance.
 */
export function createTeamsAdapter(config: TeamsAdapterConfig): TeamsAdapterImpl {
  return new TeamsAdapterImpl(config);
}

// ---------------------------------------------------------------------------
// Re-exports
// ---------------------------------------------------------------------------

export { adaptiveCardToAttachment };
export * from './types.js';
export { TeamsAdapterError, TeamsAdapterErrorCode } from './utils/error-handler.js';
export { encodeThreadId, decodeThreadId, channelIdFromThreadId, isDM as isDirectMessage } from './utils/thread-utils.js';
export { GraphClient } from './utils/graph-client.js';

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isTransient(err: TeamsAdapterError): boolean {
  return (
    err.statusCode >= 500 ||
    err.code === TeamsAdapterErrorCode.RATE_LIMITED
  );
}

function validateConfig(config: TeamsAdapterConfig): void {
  if (!config.appId) {
    throw new TeamsAdapterError(
      'TeamsAdapterConfig.appId is required',
      TeamsAdapterErrorCode.VALIDATION_ERROR,
      400
    );
  }
}

/**
 * Parses a composite "teamId:channelId" string into [teamId, channelId].
 * If no colon separator is found, the whole string is treated as the channelId
 * with an empty teamId.
 */
function parseCompositeChannelId(id: string): [string, string] {
  const idx = id.indexOf(':');
  if (idx === -1) return ['', id];
  return [id.slice(0, idx), id.slice(idx + 1)];
}

/**
 * Strips HTML tags from a string.
 */
function stripHtml(html: string): string {
  return html.replace(/<[^>]+>/g, '').trim();
}

// Graph API response shapes (internal)
interface GraphMessage {
  id?: string;
  replyToId?: string | null;
  etag?: string;
  messageType?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  body?: { content?: string; contentType?: string };
  from?: {
    user?: { id?: string; displayName?: string };
    application?: { id?: string; displayName?: string };
  };
}

interface GraphChannel {
  id?: string;
  displayName?: string;
  description?: string;
  membershipType?: string;
  createdDateTime?: string;
}

function buildGraphMessageBody(message: Message): Record<string, unknown> {
  return {
    body: {
      contentType: message.formattedContent?.type === 'html' ? 'html' : 'text',
      content: message.formattedContent
        ? renderMarkdown(message.formattedContent)
        : (message.text ?? ''),
    },
  };
}
