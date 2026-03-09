import type {
  TeamsActivity,
  TeamsAttachment,
  Message,
  FormattedContent,
  Attachment,
  Block,
} from '../types.js';

// ---------------------------------------------------------------------------
// Activity → Message
// ---------------------------------------------------------------------------

/**
 * Converts a raw Bot Framework TeamsActivity into a normalised Message object.
 */
export function activityToMessage(activity: TeamsActivity): Message {
  const message: Message = {
    id: activity.id,
    text: stripMentions(activity.text ?? ''),
    userId: activity.from.id,
    userName: activity.from.name,
    timestamp: activity.timestamp,
    replyToId: activity.replyToId,
    threadId: activity.conversation.id,
    metadata: {
      channelId: activity.channelId,
      conversationType: activity.conversation.conversationType,
      tenantId: activity.conversation.tenantId,
      serviceUrl: activity.serviceUrl,
      channelData: activity.channelData,
    },
  };

  // Convert Teams attachments
  if (activity.attachments && activity.attachments.length > 0) {
    message.attachments = activity.attachments.map(teamsAttachmentToAttachment);
  }

  return message;
}

/**
 * Converts a normalised Message into a partial TeamsActivity payload
 * suitable for sending via the Bot Connector API.
 */
export function messageToActivity(message: Message): Partial<TeamsActivity> {
  const activity: Partial<TeamsActivity> = {
    type: 'message',
  };

  if (message.formattedContent) {
    const rendered = renderMarkdown(message.formattedContent);
    activity.text = rendered;
    activity.textFormat = message.formattedContent.type === 'html' ? 'html' : 'markdown';
  } else if (message.text) {
    activity.text = message.text;
    activity.textFormat = 'markdown';
  }

  if (message.attachments && message.attachments.length > 0) {
    activity.attachments = message.attachments.map(attachmentToTeamsAttachment);
  }

  if (message.replyToId) {
    activity.replyToId = message.replyToId;
  }

  return activity;
}

// ---------------------------------------------------------------------------
// Formatted content rendering
// ---------------------------------------------------------------------------

/**
 * Renders a FormattedContent object to a Teams-compatible HTML/Markdown string.
 */
export function renderMarkdown(content: FormattedContent): string {
  switch (content.type) {
    case 'text':
      return typeof content.content === 'string' ? escapeHtml(content.content) : '';

    case 'markdown':
      return typeof content.content === 'string'
        ? markdownToTeamsHtml(content.content)
        : '';

    case 'html':
      return typeof content.content === 'string' ? content.content : '';

    case 'adaptive_card':
      // Adaptive cards are sent as attachments, not text; return empty string
      return '';

    case 'blocks':
      return blocksToHtml(content.blocks ?? []);

    default:
      return typeof content.content === 'string' ? content.content : '';
  }
}

/**
 * Wraps an Adaptive Card object in a TeamsAttachment envelope.
 */
export function adaptiveCardToAttachment(card: Record<string, unknown>): TeamsAttachment {
  return {
    contentType: 'application/vnd.microsoft.card.adaptive',
    content: {
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.4',
      ...card,
    },
  };
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Strips @mention XML tags from Teams message text.
 */
function stripMentions(text: string): string {
  return text
    .replace(/<at>[^<]*<\/at>/gi, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

/**
 * Escapes HTML special characters.
 */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}

/**
 * Converts a subset of Markdown to Teams-compatible HTML.
 * Supports: bold, italic, strikethrough, inline code, code blocks,
 * unordered/ordered lists, blockquotes, links, and images.
 */
function markdownToTeamsHtml(md: string): string {
  let html = md;

  // Fenced code blocks (``` ... ```)
  html = html.replace(/```(\w*)\n?([\s\S]*?)```/g, (_m, __lang: string, code: string) => {
    return `<pre><code>${escapeHtml(code.trim())}</code></pre>`;
  });

  // Inline code (`...`)
  html = html.replace(/`([^`]+)`/g, (_m, code: string) => `<code>${escapeHtml(code)}</code>`);

  // Bold (**text** or __text__)
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/__(.+?)__/g, '<strong>$1</strong>');

  // Italic (*text* or _text_)
  html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');
  html = html.replace(/_(.+?)_/g, '<em>$1</em>');

  // Strikethrough (~~text~~)
  html = html.replace(/~~(.+?)~~/g, '<del>$1</del>');

  // Images (![alt](url))
  html = html.replace(
    /!\[([^\]]*)\]\(([^)]+)\)/g,
    (_m, alt: string, url: string) => `<img src="${url}" alt="${alt}" />`
  );

  // Links ([text](url))
  html = html.replace(
    /\[([^\]]+)\]\(([^)]+)\)/g,
    (_m, text: string, url: string) => `<a href="${url}">${text}</a>`
  );

  // Blockquotes (> text)
  html = html.replace(/^> (.+)$/gm, '<blockquote>$1</blockquote>');

  // Unordered lists (- item or * item)
  html = html.replace(/^[-*] (.+)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>\n?)+/g, (m) => `<ul>${m}</ul>`);

  // Ordered lists (1. item)
  html = html.replace(/^\d+\. (.+)$/gm, '<li>$1</li>');

  // Line breaks
  html = html.replace(/\n/g, '<br/>');

  return html;
}

/**
 * Converts a Block array (Slack-style blocks) to HTML.
 */
function blocksToHtml(blocks: Block[]): string {
  return blocks
    .map((block) => {
      switch (block.type) {
        case 'section':
          return `<p>${block.text ?? ''}</p>`;
        case 'divider':
          return '<hr/>';
        case 'header':
          return `<h2>${block.text ?? ''}</h2>`;
        case 'context': {
          const elements = Array.isArray(block.elements) ? block.elements : [];
          const texts = elements.map((el) =>
            typeof el === 'object' && el !== null && 'text' in el
              ? String((el as Record<string, unknown>)['text'])
              : ''
          );
          return `<small>${texts.join(' | ')}</small>`;
        }
        default:
          return block.text ? `<p>${block.text}</p>` : '';
      }
    })
    .join('\n');
}

/**
 * Maps a TeamsAttachment to a normalised Attachment.
 */
function teamsAttachmentToAttachment(ta: TeamsAttachment): Attachment {
  if (
    ta.contentType === 'application/vnd.microsoft.card.adaptive' ||
    ta.contentType.includes('card')
  ) {
    return { type: 'card', content: ta.content, contentType: ta.contentType, name: ta.name };
  }
  if (ta.contentType.startsWith('image/')) {
    return { type: 'image', url: ta.contentUrl, name: ta.name, contentType: ta.contentType };
  }
  return { type: 'file', url: ta.contentUrl, name: ta.name, contentType: ta.contentType };
}

/**
 * Maps a normalised Attachment back to a TeamsAttachment.
 */
function attachmentToTeamsAttachment(att: Attachment): TeamsAttachment {
  switch (att.type) {
    case 'card':
      return {
        contentType: att.contentType ?? 'application/vnd.microsoft.card.adaptive',
        content: att.content,
        name: att.name,
      };
    case 'image':
      return {
        contentType: att.contentType ?? 'image/png',
        contentUrl: att.url,
        name: att.name,
      };
    case 'embed':
      return {
        contentType: 'application/vnd.microsoft.card.thumbnail',
        content: att.content,
        name: att.name,
      };
    default:
      return {
        contentType: att.contentType ?? 'application/octet-stream',
        contentUrl: att.url,
        name: att.name,
      };
  }
}
