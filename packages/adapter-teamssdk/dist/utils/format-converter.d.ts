import type { TeamsActivity, TeamsAttachment, Message, FormattedContent } from '../types.js';
/**
 * Converts a raw Bot Framework TeamsActivity into a normalised Message object.
 */
export declare function activityToMessage(activity: TeamsActivity): Message;
/**
 * Converts a normalised Message into a partial TeamsActivity payload
 * suitable for sending via the Bot Connector API.
 */
export declare function messageToActivity(message: Message): Partial<TeamsActivity>;
/**
 * Renders a FormattedContent object to a Teams-compatible HTML/Markdown string.
 */
export declare function renderMarkdown(content: FormattedContent): string;
/**
 * Wraps an Adaptive Card object in a TeamsAttachment envelope.
 */
export declare function adaptiveCardToAttachment(card: Record<string, unknown>): TeamsAttachment;
//# sourceMappingURL=format-converter.d.ts.map