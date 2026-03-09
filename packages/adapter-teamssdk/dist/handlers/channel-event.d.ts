import type { TeamsActivity } from '../types.js';
import type { TeamsAdapterImpl } from '../index.js';
/**
 * Creates channel- and team-level event handler functions.
 */
export declare function createChannelEventHandlers(adapter: TeamsAdapterImpl): {
    /**
     * Handles member(s) added to a conversation or team.
     * Emits 'member_added' on the ChatInstance.
     */
    onMemberAdded: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles member(s) removed from a conversation or team.
     * Emits 'member_removed' on the ChatInstance.
     */
    onMemberRemoved: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a team rename event.
     * Emits 'team_renamed' on the ChatInstance.
     */
    onTeamRenamed: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a channel created event.
     * Emits 'channel_created' on the ChatInstance.
     */
    onChannelCreated: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a channel rename event.
     * Emits 'channel_renamed' on the ChatInstance.
     */
    onChannelRenamed: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a channel deletion event.
     * Emits 'channel_deleted' on the ChatInstance.
     */
    onChannelDeleted: (activity: TeamsActivity) => Promise<void>;
};
//# sourceMappingURL=channel-event.d.ts.map