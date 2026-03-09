import type { TeamsActivity } from '../types.js';
import type { TeamsAdapterImpl } from '../index.js';
import { activityToMessage } from '../utils/format-converter.js';

/**
 * Creates message-related event handler functions.
 */
export function createMessageHandlers(adapter: TeamsAdapterImpl) {
  return {
    /**
     * Handles an incoming message activity.
     * Emits 'message' on the ChatInstance.
     */
    onMessage: async (activity: TeamsActivity): Promise<void> => {
      const message = activityToMessage(activity);
      adapter.chatInstance?.emit('message', {
        message,
        activity,
        adapter,
      });
    },

    /**
     * Handles an activity where the bot was @mentioned.
     * Emits 'mention' on the ChatInstance.
     */
    onMention: async (activity: TeamsActivity): Promise<void> => {
      const message = activityToMessage(activity);
      adapter.chatInstance?.emit('mention', {
        message,
        activity,
        adapter,
        mentionText: activity.text ?? '',
      });
    },

    /**
     * Handles a reply added to a thread.
     * Emits 'thread_reply' on the ChatInstance.
     */
    onThreadReplyAdded: async (activity: TeamsActivity): Promise<void> => {
      const message = activityToMessage(activity);
      adapter.chatInstance?.emit('thread_reply', {
        message,
        activity,
        adapter,
        replyToId: activity.replyToId,
      });
    },

    /**
     * Handles a direct message received by the bot.
     * Emits 'dm_received' on the ChatInstance.
     */
    onDMReceived: async (activity: TeamsActivity): Promise<void> => {
      const message = activityToMessage(activity);
      adapter.chatInstance?.emit('dm_received', {
        message,
        activity,
        adapter,
      });
    },

    /**
     * Handles a reaction being added to a message.
     * Emits 'reaction_added' on the ChatInstance.
     */
    onReactionAdded: async (activity: TeamsActivity): Promise<void> => {
      adapter.chatInstance?.emit('reaction_added', {
        messageId: activity.replyToId ?? activity.id,
        reactions: activity.reactionsAdded ?? [],
        userId: activity.from.id,
        userName: activity.from.name,
        activity,
        adapter,
      });
    },

    /**
     * Handles a reaction being removed from a message.
     * Emits 'reaction_removed' on the ChatInstance.
     */
    onReactionRemoved: async (activity: TeamsActivity): Promise<void> => {
      adapter.chatInstance?.emit('reaction_removed', {
        messageId: activity.replyToId ?? activity.id,
        reactions: activity.reactionsRemoved ?? [],
        userId: activity.from.id,
        userName: activity.from.name,
        activity,
        adapter,
      });
    },
  };
}
