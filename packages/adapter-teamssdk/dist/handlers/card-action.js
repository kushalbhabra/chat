/**
 * Creates card-action and invoke event handler functions.
 */
export function createCardActionHandlers(adapter) {
    return {
        /**
         * Handles a card action submit (Action.Submit on an Adaptive Card).
         * Emits 'card_action' on the ChatInstance.
         */
        onCardAction: async (activity) => {
            adapter.chatInstance?.emit('card_action', {
                actionData: activity.value,
                userId: activity.from.id,
                userName: activity.from.name,
                conversationId: activity.conversation.id,
                activity,
                adapter,
            });
        },
        /**
         * Handles a generic invoke activity (task/fetch, composeExtension, etc.).
         * Emits 'invoke' on the ChatInstance.
         */
        onInvoke: async (activity) => {
            adapter.chatInstance?.emit('invoke', {
                name: activity.name,
                value: activity.value,
                userId: activity.from.id,
                userName: activity.from.name,
                conversationId: activity.conversation.id,
                activity,
                adapter,
            });
        },
        /**
         * Handles a message action (context menu action on a message).
         * Emits 'message_action' on the ChatInstance.
         */
        onMessageAction: async (activity) => {
            adapter.chatInstance?.emit('message_action', {
                name: activity.name,
                value: activity.value,
                userId: activity.from.id,
                userName: activity.from.name,
                conversationId: activity.conversation.id,
                activity,
                adapter,
            });
        },
    };
}
//# sourceMappingURL=card-action.js.map