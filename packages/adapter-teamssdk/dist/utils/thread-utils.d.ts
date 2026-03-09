import type { TeamsContext } from '../types.js';
/**
 * Encodes a TeamsContext into a URL-safe base64 thread ID string.
 */
export declare function encodeThreadId(context: TeamsContext): string;
/**
 * Decodes a base64 thread ID string back into a TeamsContext.
 */
export declare function decodeThreadId(threadId: string): TeamsContext;
/**
 * Extracts the channel ID from an encoded thread ID.
 * Falls back to the conversationId if no explicit channelId is present.
 */
export declare function channelIdFromThreadId(threadId: string): string;
/**
 * Returns true if the thread ID represents a direct-message (DM) conversation.
 */
export declare function isDM(threadId: string): boolean;
//# sourceMappingURL=thread-utils.d.ts.map