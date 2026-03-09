/**
 * Lightweight Microsoft Graph API client.
 * All requests are authenticated with a Bearer access token.
 */
export declare class GraphClient {
    private readonly accessToken;
    private readonly baseUrl;
    constructor(accessToken: string, baseUrl?: string);
    /**
     * Lists messages in a Teams channel.
     * GET /teams/{teamId}/channels/{channelId}/messages
     */
    getChannelMessages(teamId: string, channelId: string): Promise<unknown[]>;
    /**
     * Gets a single message from a Teams channel.
     * GET /teams/{teamId}/channels/{channelId}/messages/{messageId}
     */
    getChannelMessage(teamId: string, channelId: string, messageId: string): Promise<unknown>;
    /**
     * Lists members of a Teams channel.
     * GET /teams/{teamId}/channels/{channelId}/members
     */
    getConversationMembers(teamId: string, channelId: string): Promise<unknown[]>;
    /**
     * Lists channels in a team.
     * GET /teams/{teamId}/channels
     */
    listChannels(teamId: string): Promise<unknown[]>;
    /**
     * Gets details of a specific channel.
     * GET /teams/{teamId}/channels/{channelId}
     */
    getChannel(teamId: string, channelId: string): Promise<unknown>;
    /**
     * Sets a reaction on a channel message.
     * POST /teams/{teamId}/channels/{channelId}/messages/{messageId}/setReaction
     */
    setReaction(teamId: string, channelId: string, messageId: string, reactionType: string): Promise<void>;
    /**
     * Unsets a reaction on a channel message.
     * POST /teams/{teamId}/channels/{channelId}/messages/{messageId}/unsetReaction
     */
    unsetReaction(teamId: string, channelId: string, messageId: string, reactionType: string): Promise<void>;
    /**
     * Creates a new one-on-one chat with a user.
     * POST /chats
     */
    createOneOnOneChat(userId: string, botId: string): Promise<unknown>;
    /**
     * Sends a message to a chat.
     * POST /chats/{chatId}/messages
     */
    sendChatMessage(chatId: string, body: unknown): Promise<unknown>;
    private get;
    private getList;
    private post;
    private authHeaders;
}
//# sourceMappingURL=graph-client.d.ts.map