import type { TeamsActivity } from '../types.js';
import type { TeamsAdapterImpl } from '../index.js';
/**
 * Creates message-related event handler functions.
 */
export declare function createMessageHandlers(adapter: TeamsAdapterImpl): {
    /**
     * Handles an incoming message activity.
     * Emits 'message' on the ChatInstance.
     */
    onMessage: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles an activity where the bot was @mentioned.
     * Emits 'mention' on the ChatInstance.
     */
    onMention: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a reply added to a thread.
     * Emits 'thread_reply' on the ChatInstance.
     */
    onThreadReplyAdded: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a direct message received by the bot.
     * Emits 'dm_received' on the ChatInstance.
     */
    onDMReceived: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a reaction being added to a message.
     * Emits 'reaction_added' on the ChatInstance.
     */
    onReactionAdded: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a reaction being removed from a message.
     * Emits 'reaction_removed' on the ChatInstance.
     */
    onReactionRemoved: (activity: TeamsActivity) => Promise<void>;
};
//# sourceMappingURL=message.d.ts.map