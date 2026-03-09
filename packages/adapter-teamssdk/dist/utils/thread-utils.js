/**
 * Encodes a TeamsContext into a URL-safe base64 thread ID string.
 */
export function encodeThreadId(context) {
    const json = JSON.stringify(context);
    return Buffer.from(json, 'utf8').toString('base64url');
}
/**
 * Decodes a base64 thread ID string back into a TeamsContext.
 */
export function decodeThreadId(threadId) {
    try {
        const json = Buffer.from(threadId, 'base64url').toString('utf8');
        const parsed = JSON.parse(json);
        if (typeof parsed !== 'object' ||
            parsed === null ||
            !('serviceUrl' in parsed) ||
            !('tenantId' in parsed) ||
            !('conversationId' in parsed)) {
            throw new Error('Invalid thread ID format: missing required fields');
        }
        return parsed;
    }
    catch (err) {
        throw new Error(`Failed to decode thread ID: ${err instanceof Error ? err.message : String(err)}`);
    }
}
/**
 * Extracts the channel ID from an encoded thread ID.
 * Falls back to the conversationId if no explicit channelId is present.
 */
export function channelIdFromThreadId(threadId) {
    const context = decodeThreadId(threadId);
    return context.channelId ?? context.conversationId;
}
/**
 * Returns true if the thread ID represents a direct-message (DM) conversation.
 */
export function isDM(threadId) {
    const context = decodeThreadId(threadId);
    // A DM has isDM explicitly set, or lacks both channelId and teamId
    if (context.isDM !== undefined) {
        return context.isDM;
    }
    return context.channelId === undefined && context.teamId === undefined;
}
//# sourceMappingURL=thread-utils.js.map