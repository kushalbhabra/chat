import type { TeamsActivity } from '../types.js';
import type { TeamsAdapterImpl } from '../index.js';

/**
 * Creates channel- and team-level event handler functions.
 */
export function createChannelEventHandlers(adapter: TeamsAdapterImpl) {
  return {
    /**
     * Handles member(s) added to a conversation or team.
     * Emits 'member_added' on the ChatInstance.
     */
    onMemberAdded: async (activity: TeamsActivity): Promise<void> => {
      adapter.chatInstance?.emit('member_added', {
        membersAdded: activity.membersAdded ?? [],
        conversationId: activity.conversation.id,
        teamId: extractTeamId(activity),
        channelId: extractChannelId(activity),
        activity,
        adapter,
      });
    },

    /**
     * Handles member(s) removed from a conversation or team.
     * Emits 'member_removed' on the ChatInstance.
     */
    onMemberRemoved: async (activity: TeamsActivity): Promise<void> => {
      adapter.chatInstance?.emit('member_removed', {
        membersRemoved: activity.membersRemoved ?? [],
        conversationId: activity.conversation.id,
        teamId: extractTeamId(activity),
        channelId: extractChannelId(activity),
        activity,
        adapter,
      });
    },

    /**
     * Handles a team rename event.
     * Emits 'team_renamed' on the ChatInstance.
     */
    onTeamRenamed: async (activity: TeamsActivity): Promise<void> => {
      const channelData = activity.channelData as Record<string, unknown> | undefined;
      adapter.chatInstance?.emit('team_renamed', {
        team: channelData?.['team'],
        newName: (channelData?.['team'] as Record<string, unknown> | undefined)?.['name'],
        conversationId: activity.conversation.id,
        activity,
        adapter,
      });
    },

    /**
     * Handles a channel created event.
     * Emits 'channel_created' on the ChatInstance.
     */
    onChannelCreated: async (activity: TeamsActivity): Promise<void> => {
      const channelData = activity.channelData as Record<string, unknown> | undefined;
      adapter.chatInstance?.emit('channel_created', {
        channel: channelData?.['channel'],
        teamId: extractTeamId(activity),
        activity,
        adapter,
      });
    },

    /**
     * Handles a channel rename event.
     * Emits 'channel_renamed' on the ChatInstance.
     */
    onChannelRenamed: async (activity: TeamsActivity): Promise<void> => {
      const channelData = activity.channelData as Record<string, unknown> | undefined;
      adapter.chatInstance?.emit('channel_renamed', {
        channel: channelData?.['channel'],
        teamId: extractTeamId(activity),
        activity,
        adapter,
      });
    },

    /**
     * Handles a channel deletion event.
     * Emits 'channel_deleted' on the ChatInstance.
     */
    onChannelDeleted: async (activity: TeamsActivity): Promise<void> => {
      const channelData = activity.channelData as Record<string, unknown> | undefined;
      adapter.chatInstance?.emit('channel_deleted', {
        channel: channelData?.['channel'],
        teamId: extractTeamId(activity),
        activity,
        adapter,
      });
    },
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function extractTeamId(activity: TeamsActivity): string | undefined {
  const channelData = activity.channelData as Record<string, unknown> | undefined;
  const team = channelData?.['team'] as Record<string, unknown> | undefined;
  return typeof team?.['id'] === 'string' ? team['id'] : undefined;
}

function extractChannelId(activity: TeamsActivity): string | undefined {
  const channelData = activity.channelData as Record<string, unknown> | undefined;
  const channel = channelData?.['channel'] as Record<string, unknown> | undefined;
  return typeof channel?.['id'] === 'string' ? channel['id'] : undefined;
}
