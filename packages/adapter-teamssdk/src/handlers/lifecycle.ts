import type { TeamsActivity } from '../types.js';
import type { TeamsAdapterImpl } from '../index.js';

/**
 * Creates app-lifecycle event handler functions.
 */
export function createLifecycleHandlers(adapter: TeamsAdapterImpl) {
  return {
    /**
     * Handles the bot app being installed into a team or personal scope.
     * Emits 'app_installed' on the ChatInstance.
     */
    onAppInstalled: async (activity: TeamsActivity): Promise<void> => {
      const channelData = activity.channelData as Record<string, unknown> | undefined;
      adapter.chatInstance?.emit('app_installed', {
        installedBy: activity.from.id,
        installedByName: activity.from.name,
        conversationId: activity.conversation.id,
        teamId: extractTeamId(activity),
        tenantId: activity.conversation.tenantId,
        action: channelData?.['action'],
        activity,
        adapter,
      });
    },

    /**
     * Handles the bot app being uninstalled.
     * Emits 'app_uninstalled' on the ChatInstance.
     */
    onAppUninstalled: async (activity: TeamsActivity): Promise<void> => {
      const channelData = activity.channelData as Record<string, unknown> | undefined;
      adapter.chatInstance?.emit('app_uninstalled', {
        removedBy: activity.from.id,
        removedByName: activity.from.name,
        conversationId: activity.conversation.id,
        teamId: extractTeamId(activity),
        tenantId: activity.conversation.tenantId,
        action: channelData?.['action'],
        activity,
        adapter,
      });
    },

    /**
     * Generic catch-all handler emitted for every activity received.
     * Emits 'bot_activity' on the ChatInstance.
     */
    onBotActivity: async (activity: TeamsActivity): Promise<void> => {
      adapter.chatInstance?.emit('bot_activity', {
        type: activity.type,
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
