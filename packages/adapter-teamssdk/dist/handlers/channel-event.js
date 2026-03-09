/**
 * Creates channel- and team-level event handler functions.
 */
export function createChannelEventHandlers(adapter) {
    return {
        /**
         * Handles member(s) added to a conversation or team.
         * Emits 'member_added' on the ChatInstance.
         */
        onMemberAdded: async (activity) => {
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
        onMemberRemoved: async (activity) => {
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
        onTeamRenamed: async (activity) => {
            const channelData = activity.channelData;
            adapter.chatInstance?.emit('team_renamed', {
                team: channelData?.['team'],
                newName: channelData?.['team']?.['name'],
                conversationId: activity.conversation.id,
                activity,
                adapter,
            });
        },
        /**
         * Handles a channel created event.
         * Emits 'channel_created' on the ChatInstance.
         */
        onChannelCreated: async (activity) => {
            const channelData = activity.channelData;
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
        onChannelRenamed: async (activity) => {
            const channelData = activity.channelData;
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
        onChannelDeleted: async (activity) => {
            const channelData = activity.channelData;
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
function extractTeamId(activity) {
    const channelData = activity.channelData;
    const team = channelData?.['team'];
    return typeof team?.['id'] === 'string' ? team['id'] : undefined;
}
function extractChannelId(activity) {
    const channelData = activity.channelData;
    const channel = channelData?.['channel'];
    return typeof channel?.['id'] === 'string' ? channel['id'] : undefined;
}
//# sourceMappingURL=channel-event.js.map