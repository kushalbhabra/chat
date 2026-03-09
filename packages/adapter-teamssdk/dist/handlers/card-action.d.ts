import type { TeamsActivity } from '../types.js';
import type { TeamsAdapterImpl } from '../index.js';
/**
 * Creates card-action and invoke event handler functions.
 */
export declare function createCardActionHandlers(adapter: TeamsAdapterImpl): {
    /**
     * Handles a card action submit (Action.Submit on an Adaptive Card).
     * Emits 'card_action' on the ChatInstance.
     */
    onCardAction: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a generic invoke activity (task/fetch, composeExtension, etc.).
     * Emits 'invoke' on the ChatInstance.
     */
    onInvoke: (activity: TeamsActivity) => Promise<void>;
    /**
     * Handles a message action (context menu action on a message).
     * Emits 'message_action' on the ChatInstance.
     */
    onMessageAction: (activity: TeamsActivity) => Promise<void>;
};
//# sourceMappingURL=card-action.d.ts.map