/**
 * Creates app-lifecycle event handler functions.
 */
export function createLifecycleHandlers(adapter) {
    return {
        /**
         * Handles the bot app being installed into a team or personal scope.
         * Emits 'app_installed' on the ChatInstance.
         */
        onAppInstalled: async (activity) => {
            const channelData = activity.channelData;
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
        onAppUninstalled: async (activity) => {
            const channelData = activity.channelData;
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
        onBotActivity: async (activity) => {
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
function extractTeamId(activity) {
    const channelData = activity.channelData;
    const team = channelData?.['team'];
    return typeof team?.['id'] === 'string' ? team['id'] : undefined;
}
//# sourceMappingURL=lifecycle.js.map