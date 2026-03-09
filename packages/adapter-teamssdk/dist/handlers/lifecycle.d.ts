import type { TeamsActivity } from '../types.js';
import type { TeamsAdapterImpl } from '../index.js';
/**
 * Creates app-lifecycle event handler functions.
 */
export declare function createLifecycleHandlers(adapter: TeamsAdapterImpl): {
    /**
     * Handles the bot app being installed into a team or personal scope.
     * Emits 'app_installed' on the ChatInstance.
     */
    onAppInstalled: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles the bot app being uninstalled.
     * Emits 'app_uninstalled' on the ChatInstance.
     */
    onAppUninstalled: (activity: TeamsActivity) => Promise<void>;
    /**
     * Generic catch-all handler emitted for every activity received.
     * Emits 'bot_activity' on the ChatInstance.
     */
    onBotActivity: (activity: TeamsActivity) => Promise<void>;
};
//# sourceMappingURL=lifecycle.d.ts.map