import { handleHttpError, TeamsAdapterError, TeamsAdapterErrorCode } from './error-handler.js';
const DEFAULT_GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0';
/**
 * Lightweight Microsoft Graph API client.
 * All requests are authenticated with a Bearer access token.
 */
export class GraphClient {
    accessToken;
    baseUrl;
    constructor(accessToken, baseUrl = DEFAULT_GRAPH_API_BASE) {
        this.accessToken = accessToken;
        this.baseUrl = baseUrl.replace(/\/$/, '');
    }
    // ---------------------------------------------------------------------------
    // Messages
    // ---------------------------------------------------------------------------
    /**
     * Lists messages in a Teams channel.
     * GET /teams/{teamId}/channels/{channelId}/messages
     */
    async getChannelMessages(teamId, channelId) {
        const url = `${this.baseUrl}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;
        return this.getList(url);
    }
    /**
     * Gets a single message from a Teams channel.
     * GET /teams/{teamId}/channels/{channelId}/messages/{messageId}
     */
    async getChannelMessage(teamId, channelId, messageId) {
        const url = `${this.baseUrl}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}`;
        return this.get(url);
    }
    // ---------------------------------------------------------------------------
    // Members
    // ---------------------------------------------------------------------------
    /**
     * Lists members of a Teams channel.
     * GET /teams/{teamId}/channels/{channelId}/members
     */
    async getConversationMembers(teamId, channelId) {
        const url = `${this.baseUrl}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/members`;
        return this.getList(url);
    }
    // ---------------------------------------------------------------------------
    // Channels
    // ---------------------------------------------------------------------------
    /**
     * Lists channels in a team.
     * GET /teams/{teamId}/channels
     */
    async listChannels(teamId) {
        const url = `${this.baseUrl}/teams/${encodeURIComponent(teamId)}/channels`;
        return this.getList(url);
    }
    /**
     * Gets details of a specific channel.
     * GET /teams/{teamId}/channels/{channelId}
     */
    async getChannel(teamId, channelId) {
        const url = `${this.baseUrl}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}`;
        return this.get(url);
    }
    // ---------------------------------------------------------------------------
    // Reactions
    // ---------------------------------------------------------------------------
    /**
     * Sets a reaction on a channel message.
     * POST /teams/{teamId}/channels/{channelId}/messages/{messageId}/setReaction
     */
    async setReaction(teamId, channelId, messageId, reactionType) {
        const url = `${this.baseUrl}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/setReaction`;
        await this.post(url, { reactionType });
    }
    /**
     * Unsets a reaction on a channel message.
     * POST /teams/{teamId}/channels/{channelId}/messages/{messageId}/unsetReaction
     */
    async unsetReaction(teamId, channelId, messageId, reactionType) {
        const url = `${this.baseUrl}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/unsetReaction`;
        await this.post(url, { reactionType });
    }
    // ---------------------------------------------------------------------------
    // Chat (DM / group chat)
    // ---------------------------------------------------------------------------
    /**
     * Creates a new one-on-one chat with a user.
     * POST /chats
     */
    async createOneOnOneChat(userId, botId) {
        const url = `${this.baseUrl}/chats`;
        return this.post(url, {
            chatType: 'oneOnOne',
            members: [
                {
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    roles: ['owner'],
                    'user@odata.bind': `${this.baseUrl}/users/${encodeURIComponent(userId)}`,
                },
                {
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    roles: ['owner'],
                    'user@odata.bind': `${this.baseUrl}/users/${encodeURIComponent(botId)}`,
                },
            ],
        });
    }
    /**
     * Sends a message to a chat.
     * POST /chats/{chatId}/messages
     */
    async sendChatMessage(chatId, body) {
        const url = `${this.baseUrl}/chats/${encodeURIComponent(chatId)}/messages`;
        return this.post(url, body);
    }
    // ---------------------------------------------------------------------------
    // HTTP helpers
    // ---------------------------------------------------------------------------
    async get(url) {
        const res = await fetch(url, {
            headers: this.authHeaders(),
        });
        if (!res.ok) {
            throw await handleHttpError(res);
        }
        return res.json();
    }
    async getList(url) {
        const data = (await this.get(url));
        return data.value ?? [];
    }
    async post(url, body) {
        const res = await fetch(url, {
            method: 'POST',
            headers: {
                ...this.authHeaders(),
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(body),
        });
        if (!res.ok) {
            throw await handleHttpError(res);
        }
        // 204 No Content – return empty object
        if (res.status === 204)
            return {};
        return res.json();
    }
    authHeaders() {
        if (!this.accessToken) {
            throw new TeamsAdapterError('GraphClient: missing access token', TeamsAdapterErrorCode.UNAUTHORIZED, 401);
        }
        return {
            Authorization: `Bearer ${this.accessToken}`,
            Accept: 'application/json',
        };
    }
}
//# sourceMappingURL=graph-client.js.map