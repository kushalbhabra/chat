// src/index.ts
import {
  ClientCertificateCredential,
  ClientSecretCredential,
  DefaultAzureCredential
} from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import {
  TokenCredentialAuthenticationProvider
} from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import {
  ActivityTypes,
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TeamsInfo
} from "botbuilder";
import {
  CertificateServiceClientCredentialsFactory,
  FederatedServiceClientCredentialsFactory
} from "botframework-connector";
import {
  AdapterRateLimitError,
  AuthenticationError,
  bufferToDataUri,
  extractCard,
  extractFiles,
  NetworkError,
  PermissionError,
  toBuffer,
  ValidationError
} from "@chat-adapter/shared";
import {
  ConsoleLogger,
  convertEmojiPlaceholders,
  defaultEmojiResolver,
  Message,
  NotImplementedError
} from "chat";

// src/cards.ts
import {
  createEmojiConverter,
  mapButtonStyle,
  cardToFallbackText as sharedCardToFallbackText
} from "@chat-adapter/shared";
import { cardChildToFallbackText } from "chat";
var convertEmoji = createEmojiConverter("teams");
var ADAPTIVE_CARD_SCHEMA = "http://adaptivecards.io/schemas/adaptive-card.json";
var ADAPTIVE_CARD_VERSION = "1.4";
function cardToAdaptiveCard(card) {
  const body = [];
  const actions = [];
  if (card.title) {
    body.push({
      type: "TextBlock",
      text: convertEmoji(card.title),
      weight: "bolder",
      size: "large",
      wrap: true
    });
  }
  if (card.subtitle) {
    body.push({
      type: "TextBlock",
      text: convertEmoji(card.subtitle),
      isSubtle: true,
      wrap: true
    });
  }
  if (card.imageUrl) {
    body.push({
      type: "Image",
      url: card.imageUrl,
      size: "stretch"
    });
  }
  for (const child of card.children) {
    const result = convertChildToAdaptive(child);
    body.push(...result.elements);
    actions.push(...result.actions);
  }
  const adaptiveCard = {
    type: "AdaptiveCard",
    $schema: ADAPTIVE_CARD_SCHEMA,
    version: ADAPTIVE_CARD_VERSION,
    body
  };
  if (actions.length > 0) {
    adaptiveCard.actions = actions;
  }
  return adaptiveCard;
}
function convertChildToAdaptive(child) {
  switch (child.type) {
    case "text":
      return { elements: [convertTextToElement(child)], actions: [] };
    case "image":
      return { elements: [convertImageToElement(child)], actions: [] };
    case "divider":
      return { elements: [convertDividerToElement(child)], actions: [] };
    case "actions":
      return convertActionsToElements(child);
    case "section":
      return convertSectionToElements(child);
    case "fields":
      return { elements: [convertFieldsToElement(child)], actions: [] };
    case "link":
      return {
        elements: [
          {
            type: "TextBlock",
            text: `[${convertEmoji(child.label)}](${child.url})`,
            wrap: true
          }
        ],
        actions: []
      };
    case "table":
      return { elements: [convertTableToElement(child)], actions: [] };
    default: {
      const text = cardChildToFallbackText(child);
      if (text) {
        return {
          elements: [{ type: "TextBlock", text, wrap: true }],
          actions: []
        };
      }
      return { elements: [], actions: [] };
    }
  }
}
function convertTextToElement(element) {
  const textBlock = {
    type: "TextBlock",
    text: convertEmoji(element.content),
    wrap: true
  };
  if (element.style === "bold") {
    textBlock.weight = "bolder";
  } else if (element.style === "muted") {
    textBlock.isSubtle = true;
  }
  return textBlock;
}
function convertImageToElement(element) {
  return {
    type: "Image",
    url: element.url,
    altText: element.alt || "Image",
    size: "auto"
  };
}
function convertDividerToElement(_element) {
  return {
    type: "Container",
    separator: true,
    items: []
  };
}
function convertActionsToElements(element) {
  const actions = element.children.filter((child) => child.type === "button" || child.type === "link-button").map((button) => {
    if (button.type === "link-button") {
      return convertLinkButtonToAction(button);
    }
    return convertButtonToAction(button);
  });
  return { elements: [], actions };
}
function convertButtonToAction(button) {
  const action = {
    type: "Action.Submit",
    title: convertEmoji(button.label),
    data: {
      actionId: button.id,
      value: button.value
    }
  };
  const style = mapButtonStyle(button.style, "teams");
  if (style) {
    action.style = style;
  }
  return action;
}
function convertLinkButtonToAction(button) {
  const action = {
    type: "Action.OpenUrl",
    title: convertEmoji(button.label),
    url: button.url
  };
  const style = mapButtonStyle(button.style, "teams");
  if (style) {
    action.style = style;
  }
  return action;
}
function convertSectionToElements(element) {
  const elements = [];
  const actions = [];
  const containerItems = [];
  for (const child of element.children) {
    const result = convertChildToAdaptive(child);
    containerItems.push(...result.elements);
    actions.push(...result.actions);
  }
  if (containerItems.length > 0) {
    elements.push({
      type: "Container",
      items: containerItems
    });
  }
  return { elements, actions };
}
function convertTableToElement(element) {
  const columns = element.headers.map((header) => ({
    type: "Column",
    width: "stretch",
    items: [
      {
        type: "TextBlock",
        text: convertEmoji(header),
        weight: "bolder",
        wrap: true
      }
    ]
  }));
  const headerRow = {
    type: "ColumnSet",
    columns
  };
  const dataRows = element.rows.map((row) => ({
    type: "ColumnSet",
    columns: row.map((cell) => ({
      type: "Column",
      width: "stretch",
      items: [
        {
          type: "TextBlock",
          text: convertEmoji(cell),
          wrap: true
        }
      ]
    }))
  }));
  return {
    type: "Container",
    items: [headerRow, ...dataRows]
  };
}
function convertFieldsToElement(element) {
  const facts = element.children.map((field) => ({
    title: convertEmoji(field.label),
    value: convertEmoji(field.value)
  }));
  return {
    type: "FactSet",
    facts
  };
}
function cardToFallbackText(card) {
  return sharedCardToFallbackText(card, {
    boldFormat: "**",
    lineBreak: "\n\n",
    platform: "teams"
  });
}

// src/markdown.ts
import { escapeTableCell } from "@chat-adapter/shared";
import {
  BaseFormatConverter,
  getNodeChildren,
  isBlockquoteNode,
  isCodeNode,
  isDeleteNode,
  isEmphasisNode,
  isInlineCodeNode,
  isLinkNode,
  isListNode,
  isParagraphNode,
  isStrongNode,
  isTableNode,
  isTextNode,
  parseMarkdown
} from "chat";
var TeamsSDKFormatConverter = class extends BaseFormatConverter {
  /**
   * Convert @mentions to Teams format in plain text.
   * @name → <at>name</at>
   */
  convertMentionsToTeams(text) {
    return text.replace(/@(\w+)/g, "<at>$1</at>");
  }
  /**
   * Override renderPostable to convert @mentions in plain strings.
   */
  renderPostable(message) {
    if (typeof message === "string") {
      return this.convertMentionsToTeams(message);
    }
    if ("raw" in message) {
      return this.convertMentionsToTeams(message.raw);
    }
    if ("markdown" in message) {
      return this.fromAst(parseMarkdown(message.markdown));
    }
    if ("ast" in message) {
      return this.fromAst(message.ast);
    }
    return "";
  }
  /**
   * Render an AST to Teams format.
   * Teams accepts standard markdown, so we just stringify cleanly.
   */
  fromAst(ast) {
    return this.fromAstWithNodeConverter(ast, (node) => this.nodeToTeams(node));
  }
  /**
   * Parse Teams message into an AST.
   * Converts Teams HTML/mentions to standard markdown format.
   */
  toAst(teamsText) {
    let markdown = teamsText;
    markdown = markdown.replace(/<at>([^<]+)<\/at>/gi, "@$1");
    markdown = markdown.replace(
      /<(b|strong)>([^<]+)<\/(b|strong)>/gi,
      "**$2**"
    );
    markdown = markdown.replace(/<(i|em)>([^<]+)<\/(i|em)>/gi, "_$2_");
    markdown = markdown.replace(
      /<(s|strike)>([^<]+)<\/(s|strike)>/gi,
      "~~$2~~"
    );
    markdown = markdown.replace(
      /<a[^>]+href="([^"]+)"[^>]*>([^<]+)<\/a>/gi,
      "[$2]($1)"
    );
    markdown = markdown.replace(/<code>([^<]+)<\/code>/gi, "`$1`");
    markdown = markdown.replace(/<pre>([^<]+)<\/pre>/gi, "```\n$1\n```");
    let prev;
    do {
      prev = markdown;
      markdown = markdown.replace(/<[^>]+>/g, "");
    } while (markdown !== prev);
    const entityMap = {
      "&lt;": "<",
      "&gt;": ">",
      "&amp;": "&",
      "&quot;": '"',
      "&#39;": "'"
    };
    markdown = markdown.replace(
      /&(?:lt|gt|amp|quot|#39);/g,
      (match) => entityMap[match] ?? match
    );
    return parseMarkdown(markdown);
  }
  nodeToTeams(node) {
    if (isParagraphNode(node)) {
      return getNodeChildren(node).map((child) => this.nodeToTeams(child)).join("");
    }
    if (isTextNode(node)) {
      return node.value.replace(/@(\w+)/g, "<at>$1</at>");
    }
    if (isStrongNode(node)) {
      const content = getNodeChildren(node).map((child) => this.nodeToTeams(child)).join("");
      return `**${content}**`;
    }
    if (isEmphasisNode(node)) {
      const content = getNodeChildren(node).map((child) => this.nodeToTeams(child)).join("");
      return `_${content}_`;
    }
    if (isDeleteNode(node)) {
      const content = getNodeChildren(node).map((child) => this.nodeToTeams(child)).join("");
      return `~~${content}~~`;
    }
    if (isInlineCodeNode(node)) {
      return `\`${node.value}\``;
    }
    if (isCodeNode(node)) {
      return `\`\`\`${node.lang || ""}
${node.value}
\`\`\``;
    }
    if (isLinkNode(node)) {
      const linkText = getNodeChildren(node).map((child) => this.nodeToTeams(child)).join("");
      return `[${linkText}](${node.url})`;
    }
    if (isBlockquoteNode(node)) {
      return getNodeChildren(node).map((child) => `> ${this.nodeToTeams(child)}`).join("\n");
    }
    if (isListNode(node)) {
      return this.renderList(node, 0, (child) => this.nodeToTeams(child));
    }
    if (node.type === "break") {
      return "\n";
    }
    if (node.type === "thematicBreak") {
      return "---";
    }
    if (isTableNode(node)) {
      return this.tableToGfm(node);
    }
    return this.defaultNodeToText(node, (child) => this.nodeToTeams(child));
  }
  /**
   * Render an mdast table node as a GFM markdown table.
   * Teams renders markdown tables natively.
   */
  tableToGfm(node) {
    const rows = [];
    for (const row of node.children) {
      const cells = [];
      for (const cell of row.children) {
        const cellContent = getNodeChildren(cell).map((child) => this.nodeToTeams(child)).join("");
        cells.push(cellContent);
      }
      rows.push(cells);
    }
    if (rows.length === 0) {
      return "";
    }
    const lines = [];
    lines.push(`| ${rows[0].map(escapeTableCell).join(" | ")} |`);
    const separators = rows[0].map(() => "---");
    lines.push(`| ${separators.join(" | ")} |`);
    for (let i = 1; i < rows.length; i++) {
      lines.push(`| ${rows[i].map(escapeTableCell).join(" | ")} |`);
    }
    return lines.join("\n");
  }
};

// src/index.ts
var MESSAGEID_CAPTURE_PATTERN = /messageid=(\d+)/;
var MESSAGEID_STRIP_PATTERN = /;messageid=\d+/;
var SEMICOLON_MESSAGEID_CAPTURE_PATTERN = /;messageid=(\d+)/;
var TeamsApp = class {
  handlers = /* @__PURE__ */ new Map();
  // -- Registration methods --
  /** Handle regular channel/group messages */
  $onMessage(handler) {
    return this.addHandler("message", handler);
  }
  /** Handle @mention messages (the bot was explicitly mentioned) */
  $onMention(handler) {
    return this.addHandler("mention", handler);
  }
  /** Handle replies in a thread */
  $onThreadReplyAdded(handler) {
    return this.addHandler("threadReply", handler);
  }
  /** Handle 1:1 direct messages sent to the bot */
  $onDMReceived(handler) {
    return this.addHandler("dm", handler);
  }
  /** Handle emoji reactions added to a message */
  $onReactionAdded(handler) {
    return this.addHandler("reactionAdded", handler);
  }
  /** Handle emoji reactions removed from a message */
  $onReactionRemoved(handler) {
    return this.addHandler("reactionRemoved", handler);
  }
  /** Handle Adaptive Card Action.Submit clicks */
  $onCardAction(handler) {
    return this.addHandler("cardAction", handler);
  }
  /** Handle invoke activities (task modules, adaptive card invokes) */
  $onInvoke(handler) {
    return this.addHandler("invoke", handler);
  }
  /** Handle message context-menu actions */
  $onMessageAction(handler) {
    return this.addHandler("messageAction", handler);
  }
  /** Handle member added to team/conversation */
  $onMemberAdded(handler) {
    return this.addHandler("memberAdded", handler);
  }
  /** Handle member removed from team/conversation */
  $onMemberRemoved(handler) {
    return this.addHandler("memberRemoved", handler);
  }
  /** Handle team rename event */
  $onTeamRenamed(handler) {
    return this.addHandler("teamRenamed", handler);
  }
  /** Handle channel created in a team */
  $onChannelCreated(handler) {
    return this.addHandler("channelCreated", handler);
  }
  /** Handle channel renamed in a team */
  $onChannelRenamed(handler) {
    return this.addHandler("channelRenamed", handler);
  }
  /** Handle channel deleted from a team */
  $onChannelDeleted(handler) {
    return this.addHandler("channelDeleted", handler);
  }
  /** Handle app installation in a team/personal scope */
  $onAppInstalled(handler) {
    return this.addHandler("appInstalled", handler);
  }
  /** Handle app uninstallation */
  $onAppUninstalled(handler) {
    return this.addHandler("appUninstalled", handler);
  }
  /** Handle any bot activity (fires for every activity, regardless of type) */
  $onBotActivity(handler) {
    return this.addHandler("botActivity", handler);
  }
  // -- Dispatch --
  /**
   * Route an incoming activity to all matching registered handlers.
   * Always fires $onBotActivity handlers, then type-specific handlers.
   */
  async processActivity(activity, context) {
    await this.runHandlers("botActivity", activity, context);
    switch (activity.type) {
      case ActivityTypes.Message:
        await this.dispatchMessageActivity(activity, context);
        break;
      case ActivityTypes.MessageReaction:
        await this.dispatchReactionActivity(activity, context);
        break;
      case ActivityTypes.Invoke:
        await this.dispatchInvokeActivity(activity, context);
        break;
      case ActivityTypes.ConversationUpdate:
        await this.dispatchConversationUpdateActivity(activity, context);
        break;
      case ActivityTypes.InstallationUpdate:
        await this.dispatchInstallationUpdateActivity(activity, context);
        break;
    }
  }
  // -- Private dispatch helpers --
  async dispatchMessageActivity(activity, context) {
    const actionValue = activity.value;
    if (actionValue?.actionId) {
      await this.runHandlers("cardAction", activity, context);
      return;
    }
    const isDM = activity.conversation?.conversationType === "personal" || !activity.conversation?.isGroup;
    const isMention = (activity.entities ?? []).some(
      (e) => e.type === "mention" && e.mentioned?.id
    );
    const isReply = !!activity.replyToId;
    if (isDM) {
      await this.runHandlers("dm", activity, context);
    } else if (isMention) {
      await this.runHandlers("mention", activity, context);
    } else if (isReply) {
      await this.runHandlers("threadReply", activity, context);
    } else {
      await this.runHandlers("message", activity, context);
    }
  }
  async dispatchReactionActivity(activity, context) {
    if ((activity.reactionsAdded ?? []).length > 0) {
      await this.runHandlers("reactionAdded", activity, context);
    }
    if ((activity.reactionsRemoved ?? []).length > 0) {
      await this.runHandlers("reactionRemoved", activity, context);
    }
  }
  async dispatchInvokeActivity(activity, context) {
    if (activity.name === "adaptiveCard/action") {
      await this.runHandlers("cardAction", activity, context);
    } else {
      await this.runHandlers("invoke", activity, context);
    }
  }
  async dispatchConversationUpdateActivity(activity, context) {
    const channelData = activity.channelData;
    const eventType = channelData?.eventType;
    if ((activity.membersAdded ?? []).length > 0) {
      await this.runHandlers("memberAdded", activity, context);
    }
    if ((activity.membersRemoved ?? []).length > 0) {
      await this.runHandlers("memberRemoved", activity, context);
    }
    if (eventType === "teamRenamed") {
      await this.runHandlers("teamRenamed", activity, context);
    } else if (eventType === "channelCreated") {
      await this.runHandlers("channelCreated", activity, context);
    } else if (eventType === "channelRenamed") {
      await this.runHandlers("channelRenamed", activity, context);
    } else if (eventType === "channelDeleted") {
      await this.runHandlers("channelDeleted", activity, context);
    }
  }
  async dispatchInstallationUpdateActivity(activity, context) {
    if (activity.action === "add") {
      await this.runHandlers("appInstalled", activity, context);
    } else if (activity.action === "remove") {
      await this.runHandlers("appUninstalled", activity, context);
    }
  }
  addHandler(event, handler) {
    const existing = this.handlers.get(event) ?? [];
    existing.push(handler);
    this.handlers.set(event, existing);
    return this;
  }
  async runHandlers(event, activity, context) {
    const list = this.handlers.get(event) ?? [];
    for (const handler of list) {
      await handler(activity, context);
    }
  }
};
var ServerlessCloudAdapter = class extends CloudAdapter {
  handleActivity(authHeader, activity, logic) {
    return this.processActivity(authHeader, activity, logic);
  }
};
var TeamsSDKAdapter = class {
  name = "teamssdk";
  userName;
  botUserId;
  /** The internal TeamsApp event router (teams.ts-style API) */
  app;
  botAdapter;
  graphClient = null;
  chat = null;
  logger;
  formatConverter = new TeamsSDKFormatConverter();
  config;
  constructor(config = {}) {
    const appId = config.appId ?? process.env.TEAMS_APP_ID;
    if (!appId) {
      throw new ValidationError(
        "teamssdk",
        "appId is required. Set TEAMS_APP_ID or provide it in config."
      );
    }
    const hasExplicitAuth = config.appPassword || config.certificate || config.federated;
    const appPassword = hasExplicitAuth ? config.appPassword : config.appPassword ?? process.env.TEAMS_APP_PASSWORD;
    const appTenantId = config.appTenantId ?? process.env.TEAMS_APP_TENANT_ID;
    this.config = { ...config, appId, appPassword, appTenantId };
    this.logger = config.logger ?? new ConsoleLogger("info").child("teamssdk");
    this.userName = config.userName || "bot";
    const authMethodCount = [
      appPassword,
      config.certificate,
      config.federated
    ].filter(Boolean).length;
    if (authMethodCount === 0) {
      throw new ValidationError(
        "teamssdk",
        "One of appPassword, certificate, or federated must be provided"
      );
    }
    if (authMethodCount > 1) {
      throw new ValidationError(
        "teamssdk",
        "Only one of appPassword, certificate, or federated can be provided"
      );
    }
    if (config.appType === "SingleTenant" && !appTenantId) {
      throw new ValidationError(
        "teamssdk",
        "appTenantId is required for SingleTenant app type"
      );
    }
    const botFrameworkConfig = {
      MicrosoftAppId: appId,
      MicrosoftAppType: config.appType || "MultiTenant",
      MicrosoftAppTenantId: config.appType === "SingleTenant" ? appTenantId : void 0
    };
    let credentialsFactory;
    let graphCredential;
    if (config.certificate) {
      const { certificatePrivateKey, certificateThumbprint, x5c } = config.certificate;
      if (x5c) {
        credentialsFactory = new CertificateServiceClientCredentialsFactory(
          appId,
          x5c,
          certificatePrivateKey,
          appTenantId
        );
      } else if (certificateThumbprint) {
        credentialsFactory = new CertificateServiceClientCredentialsFactory(
          appId,
          certificateThumbprint,
          certificatePrivateKey,
          appTenantId
        );
      } else {
        throw new ValidationError(
          "teamssdk",
          "Certificate auth requires either certificateThumbprint or x5c"
        );
      }
      if (appTenantId) {
        graphCredential = new ClientCertificateCredential(appTenantId, appId, {
          certificate: certificatePrivateKey
        });
      }
    } else if (config.federated) {
      credentialsFactory = new FederatedServiceClientCredentialsFactory(
        appId,
        config.federated.clientId,
        appTenantId,
        config.federated.clientAudience
      );
      if (appTenantId) {
        graphCredential = new DefaultAzureCredential();
      }
    } else if (appPassword && appTenantId) {
      graphCredential = new ClientSecretCredential(
        appTenantId,
        appId,
        appPassword
      );
    }
    const auth = new ConfigurationBotFrameworkAuthentication(
      {
        ...botFrameworkConfig,
        ...appPassword ? { MicrosoftAppPassword: appPassword } : {}
      },
      credentialsFactory
    );
    this.botAdapter = new ServerlessCloudAdapter(auth);
    if (graphCredential) {
      const authProvider = new TokenCredentialAuthenticationProvider(
        graphCredential,
        {
          scopes: ["https://graph.microsoft.com/.default"]
        }
      );
      this.graphClient = Client.initWithMiddleware({ authProvider });
    }
    this.app = new TeamsApp();
    this.wireAppHandlers();
  }
  /**
   * Wire up the TeamsApp event handlers to forward activities to the
   * Chat instance (processMessage / processAction / processReaction).
   */
  wireAppHandlers() {
    for (const event of ["message", "mention", "threadReply", "dm"]) {
      this.app[event === "message" ? "$onMessage" : event === "mention" ? "$onMention" : event === "threadReply" ? "$onThreadReplyAdded" : "$onDMReceived"]((activity) => this.handleMessageActivity(activity));
    }
    this.app.$onReactionAdded(
      (activity) => this.handleReactionActivity(activity, true)
    );
    this.app.$onReactionRemoved(
      (activity) => this.handleReactionActivity(activity, false)
    );
    this.app.$onCardAction(
      (activity, context) => this.handleCardActionActivity(activity, context)
    );
    this.app.$onInvoke(
      (activity, context) => this.handleInvokeActivity(activity, context)
    );
  }
  async initialize(chat) {
    this.chat = chat;
  }
  async handleWebhook(request, options) {
    const body = await request.text();
    this.logger.debug("Teams SDK webhook raw body", { body });
    let activity;
    try {
      activity = JSON.parse(body);
    } catch (e) {
      this.logger.error("Failed to parse request body", { error: e });
      return new Response("Invalid JSON", { status: 400 });
    }
    const authHeader = request.headers.get("authorization") ?? "";
    try {
      await this.botAdapter.handleActivity(
        authHeader,
        activity,
        async (context) => {
          await this.handleTurn(context, options);
        }
      );
      return new Response(JSON.stringify({}), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    } catch (error) {
      this.logger.error("Bot adapter process error", { error });
      return new Response(JSON.stringify({ error: "Internal error" }), {
        status: 500,
        headers: { "Content-Type": "application/json" }
      });
    }
  }
  async handleTurn(context, _options) {
    if (!this.chat) {
      this.logger.warn("Chat instance not initialized, ignoring event");
      return;
    }
    const activity = context.activity;
    if (activity.from?.id && activity.serviceUrl) {
      const userId = activity.from.id;
      const channelData = activity.channelData;
      const tenantId = channelData?.tenant?.id;
      const ttl = 30 * 24 * 60 * 60 * 1e3;
      this.chat.getState().set(`teamssdk:serviceUrl:${userId}`, activity.serviceUrl, ttl).catch((err) => {
        this.logger.error("Failed to cache serviceUrl", { userId, error: err });
      });
      if (tenantId) {
        this.chat.getState().set(`teamssdk:tenantId:${userId}`, tenantId, ttl).catch((err) => {
          this.logger.error("Failed to cache tenantId", { userId, error: err });
        });
      }
      const team = channelData?.team;
      const teamAadGroupId = team?.aadGroupId;
      const teamThreadId = team?.id;
      const conversationId = activity.conversation?.id ?? "";
      const baseChannelId = conversationId.replace(MESSAGEID_STRIP_PATTERN, "");
      if (teamAadGroupId && channelData?.channel?.id && tenantId) {
        const ctx = {
          teamId: teamAadGroupId,
          channelId: channelData.channel.id,
          tenantId
        };
        const ctxJson = JSON.stringify(ctx);
        this.chat.getState().set(`teamssdk:channelContext:${baseChannelId}`, ctxJson, ttl).catch(() => void 0);
        if (teamThreadId) {
          this.chat.getState().set(`teamssdk:teamContext:${teamThreadId}`, ctxJson, ttl).catch(() => void 0);
        }
      } else if (teamThreadId && channelData?.channel?.id && tenantId) {
        const cachedTeamContext = await this.chat.getState().get(`teamssdk:teamContext:${teamThreadId}`);
        if (cachedTeamContext) {
          this.chat.getState().set(`teamssdk:channelContext:${baseChannelId}`, cachedTeamContext, ttl).catch(() => void 0);
        } else {
          try {
            const teamDetails = await TeamsInfo.getTeamDetails(context);
            if (teamDetails?.aadGroupId) {
              const fetchedCtx = {
                teamId: teamDetails.aadGroupId,
                channelId: channelData.channel.id,
                tenantId
              };
              const fetchedJson = JSON.stringify(fetchedCtx);
              this.chat.getState().set(`teamssdk:channelContext:${baseChannelId}`, fetchedJson, ttl).catch(() => void 0);
              this.chat.getState().set(`teamssdk:teamContext:${teamThreadId}`, fetchedJson, ttl).catch(() => void 0);
            }
          } catch {
          }
        }
      }
    }
    await this.app.processActivity(activity, context);
  }
  // ---------------------------------------------------------------------------
  // TeamsApp event implementation handlers
  // ---------------------------------------------------------------------------
  handleMessageActivity(activity, options) {
    if (!this.chat) return;
    const threadId = this.encodeThreadId({
      conversationId: activity.conversation?.id ?? "",
      serviceUrl: activity.serviceUrl ?? "",
      replyToId: activity.replyToId
    });
    this.chat.processMessage(
      this,
      threadId,
      this.parseTeamsMessage(activity, threadId),
      options
    );
  }
  handleReactionActivity(activity, added, options) {
    if (!this.chat) return;
    const conversationId = activity.conversation?.id ?? "";
    const messageIdMatch = conversationId.match(MESSAGEID_CAPTURE_PATTERN);
    const messageId = messageIdMatch?.[1] ?? activity.replyToId ?? "";
    const threadId = this.encodeThreadId({
      conversationId,
      serviceUrl: activity.serviceUrl ?? ""
    });
    const user = {
      userId: activity.from?.id ?? "unknown",
      userName: activity.from?.name ?? "unknown",
      fullName: activity.from?.name,
      isBot: false,
      isMe: this.isMessageFromSelf(activity)
    };
    const reactions = added ? activity.reactionsAdded ?? [] : activity.reactionsRemoved ?? [];
    for (const reaction of reactions) {
      const rawEmoji = reaction.type ?? "";
      const emojiValue = defaultEmojiResolver.fromTeams(rawEmoji);
      const event = {
        emoji: emojiValue,
        rawEmoji,
        added,
        user,
        messageId,
        threadId,
        raw: activity
      };
      this.chat.processReaction({ ...event, adapter: this }, options);
    }
  }
  handleCardActionActivity(activity, context, options) {
    if (!this.chat) return;
    const actionValue = activity.value;
    if (!actionValue?.actionId) return;
    const threadId = this.encodeThreadId({
      conversationId: activity.conversation?.id ?? "",
      serviceUrl: activity.serviceUrl ?? ""
    });
    const actionEvent = {
      actionId: actionValue.actionId,
      value: actionValue.value,
      user: {
        userId: activity.from?.id ?? "unknown",
        userName: activity.from?.name ?? "unknown",
        fullName: activity.from?.name ?? "unknown",
        isBot: false,
        isMe: false
      },
      messageId: activity.replyToId ?? activity.id ?? "",
      threadId,
      adapter: this,
      raw: activity
    };
    this.chat.processAction(actionEvent, options);
  }
  async handleInvokeActivity(activity, context, options) {
    if (!this.chat) return;
    if (activity.name === "adaptiveCard/action") {
      const actionData = activity.value?.action?.data;
      if (!actionData?.actionId) {
        await context.sendActivity({
          type: ActivityTypes.InvokeResponse,
          value: { status: 200 }
        });
        return;
      }
      const threadId = this.encodeThreadId({
        conversationId: activity.conversation?.id ?? "",
        serviceUrl: activity.serviceUrl ?? ""
      });
      const actionEvent = {
        actionId: actionData.actionId,
        value: actionData.value,
        user: {
          userId: activity.from?.id ?? "unknown",
          userName: activity.from?.name ?? "unknown",
          fullName: activity.from?.name ?? "unknown",
          isBot: false,
          isMe: false
        },
        messageId: activity.replyToId ?? activity.id ?? "",
        threadId,
        adapter: this,
        raw: activity
      };
      this.chat.processAction(actionEvent, options);
      await context.sendActivity({
        type: ActivityTypes.InvokeResponse,
        value: { status: 200 }
      });
    }
  }
  // ---------------------------------------------------------------------------
  // Adapter interface implementation
  // ---------------------------------------------------------------------------
  async postMessage(threadId, message) {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);
    const files = extractFiles(message);
    const fileAttachments = files.length > 0 ? await this.filesToAttachments(files) : [];
    const card = extractCard(message);
    let activity;
    if (card) {
      const adaptiveCard = cardToAdaptiveCard(card);
      activity = {
        type: ActivityTypes.Message,
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: adaptiveCard
          },
          ...fileAttachments
        ]
      };
    } else {
      const text = convertEmojiPlaceholders(
        this.formatConverter.renderPostable(message),
        "teams"
      );
      activity = {
        type: ActivityTypes.Message,
        text,
        textFormat: "markdown",
        attachments: fileAttachments.length > 0 ? fileAttachments : void 0
      };
    }
    const conversationReference = {
      channelId: "msteams",
      serviceUrl,
      conversation: { id: conversationId }
    };
    let messageId = "";
    try {
      await this.botAdapter.continueConversationAsync(
        this.config.appId,
        conversationReference,
        async (context) => {
          const response = await context.sendActivity(activity);
          messageId = response?.id ?? "";
        }
      );
    } catch (error) {
      this.handleTeamsError(error, "postMessage");
    }
    return { id: messageId, threadId, raw: activity };
  }
  async filesToAttachments(files) {
    const attachments = [];
    for (const file of files) {
      const buffer = await toBuffer(file.data, {
        platform: "teams",
        throwOnUnsupported: false
      });
      if (!buffer) continue;
      const mimeType = file.mimeType ?? "application/octet-stream";
      const dataUri = bufferToDataUri(buffer, mimeType);
      attachments.push({ contentType: mimeType, contentUrl: dataUri, name: file.filename });
    }
    return attachments;
  }
  async editMessage(threadId, messageId, message) {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);
    const card = extractCard(message);
    let activity;
    if (card) {
      const adaptiveCard = cardToAdaptiveCard(card);
      activity = {
        id: messageId,
        type: ActivityTypes.Message,
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: adaptiveCard
          }
        ]
      };
    } else {
      const text = convertEmojiPlaceholders(
        this.formatConverter.renderPostable(message),
        "teams"
      );
      activity = { id: messageId, type: ActivityTypes.Message, text, textFormat: "markdown" };
    }
    const conversationReference = {
      channelId: "msteams",
      serviceUrl,
      conversation: { id: conversationId }
    };
    try {
      await this.botAdapter.continueConversationAsync(
        this.config.appId,
        conversationReference,
        async (context) => {
          await context.updateActivity(activity);
        }
      );
    } catch (error) {
      this.handleTeamsError(error, "editMessage");
    }
    return { id: messageId, threadId, raw: activity };
  }
  async deleteMessage(threadId, messageId) {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);
    const conversationReference = {
      channelId: "msteams",
      serviceUrl,
      conversation: { id: conversationId }
    };
    try {
      await this.botAdapter.continueConversationAsync(
        this.config.appId,
        conversationReference,
        async (context) => {
          await context.deleteActivity(messageId);
        }
      );
    } catch (error) {
      this.handleTeamsError(error, "deleteMessage");
    }
  }
  async addReaction(_threadId, _messageId, _emoji) {
    throw new NotImplementedError(
      "Teams Bot Framework does not expose reaction APIs",
      "addReaction"
    );
  }
  async removeReaction(_threadId, _messageId, _emoji) {
    throw new NotImplementedError(
      "Teams Bot Framework does not expose reaction APIs",
      "removeReaction"
    );
  }
  async startTyping(threadId, _status) {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);
    const conversationReference = {
      channelId: "msteams",
      serviceUrl,
      conversation: { id: conversationId }
    };
    try {
      await this.botAdapter.continueConversationAsync(
        this.config.appId,
        conversationReference,
        async (context) => {
          await context.sendActivity({ type: ActivityTypes.Typing });
        }
      );
    } catch (error) {
      this.handleTeamsError(error, "startTyping");
    }
  }
  async openDM(userId) {
    const cachedServiceUrl = await this.chat?.getState().get(`teamssdk:serviceUrl:${userId}`);
    const cachedTenantId = await this.chat?.getState().get(`teamssdk:tenantId:${userId}`);
    const serviceUrl = cachedServiceUrl ?? "https://smba.trafficmanager.net/teams/";
    const tenantId = cachedTenantId ?? this.config.appTenantId;
    if (!tenantId) {
      throw new ValidationError(
        "teamssdk",
        "Cannot open DM: tenant ID not found. User must interact with the bot first."
      );
    }
    let conversationId = "";
    await this.botAdapter.createConversationAsync(
      this.config.appId,
      "msteams",
      serviceUrl,
      "",
      {
        isGroup: false,
        bot: { id: this.config.appId, name: this.userName },
        members: [{ id: userId }],
        tenantId,
        channelData: { tenant: { id: tenantId } }
      },
      async (turnContext) => {
        conversationId = turnContext?.activity?.conversation?.id ?? "";
      }
    );
    if (!conversationId) {
      throw new NetworkError(
        "teamssdk",
        "Failed to create 1:1 conversation - no ID returned"
      );
    }
    return this.encodeThreadId({ conversationId, serviceUrl });
  }
  async fetchMessages(threadId, options = {}) {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "fetchMessages requires appTenantId for Microsoft Graph API access.",
        "fetchMessages"
      );
    }
    const { conversationId } = this.decodeThreadId(threadId);
    const limit = options.limit ?? 50;
    const cursor = options.cursor;
    const direction = options.direction ?? "backward";
    const messageIdMatch = conversationId.match(
      SEMICOLON_MESSAGEID_CAPTURE_PATTERN
    );
    const threadMessageId = messageIdMatch?.[1];
    const baseConversationId = conversationId.replace(
      MESSAGEID_STRIP_PATTERN,
      ""
    );
    let channelContext = null;
    if (threadMessageId && this.chat) {
      const cachedContext = await this.chat.getState().get(`teamssdk:channelContext:${baseConversationId}`);
      if (cachedContext) {
        try {
          channelContext = JSON.parse(cachedContext);
        } catch {
        }
      }
    }
    try {
      if (channelContext && threadMessageId) {
        return this.fetchChannelThreadMessages(
          channelContext,
          threadMessageId,
          threadId,
          options
        );
      }
      let graphMessages;
      let hasMoreMessages = false;
      if (direction === "forward") {
        const allMessages = [];
        let nextLink;
        const apiUrl = `/chats/${encodeURIComponent(baseConversationId)}/messages`;
        do {
          const request = nextLink ? this.graphClient.api(nextLink) : this.graphClient.api(apiUrl).top(50).orderby("createdDateTime desc");
          const response = await request.get();
          allMessages.push(...response.value ?? []);
          nextLink = response["@odata.nextLink"];
        } while (nextLink);
        allMessages.reverse();
        let startIndex = 0;
        if (cursor) {
          startIndex = allMessages.findIndex(
            (msg) => msg.createdDateTime && msg.createdDateTime > cursor
          );
          if (startIndex === -1) startIndex = allMessages.length;
        }
        hasMoreMessages = startIndex + limit < allMessages.length;
        graphMessages = allMessages.slice(startIndex, startIndex + limit);
      } else {
        let request = this.graphClient.api(`/chats/${encodeURIComponent(baseConversationId)}/messages`).top(limit).orderby("createdDateTime desc");
        if (cursor) {
          request = request.filter(`createdDateTime lt ${cursor}`);
        }
        const response = await request.get();
        graphMessages = response.value ?? [];
        graphMessages.reverse();
        hasMoreMessages = graphMessages.length >= limit;
      }
      if (threadMessageId && !channelContext) {
        graphMessages = graphMessages.filter(
          (msg) => msg.id && msg.id >= threadMessageId
        );
      }
      const messages = this.graphMessagesToMessages(graphMessages, threadId);
      let nextCursor;
      if (hasMoreMessages && graphMessages.length > 0) {
        const refMsg = direction === "forward" ? graphMessages.at(-1) : graphMessages[0];
        if (refMsg?.createdDateTime) nextCursor = refMsg.createdDateTime;
      }
      return { messages, nextCursor };
    } catch (error) {
      this.logger.error("Teams Graph API: fetchMessages error", { error });
      if (error instanceof Error && error.message?.includes("403")) {
        throw new PermissionError(
          "teamssdk",
          "fetchMessages requires ChatMessage.Read.Chat or Chat.Read.All Graph permission"
        );
      }
      throw error;
    }
  }
  async fetchChannelThreadMessages(channelContext, threadMessageId, threadId, options) {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "fetchMessages requires graphClient",
        "fetchMessages"
      );
    }
    const limit = options.limit ?? 50;
    const cursor = options.cursor;
    const apiBase = `/teams/${encodeURIComponent(channelContext.teamId)}/channels/${encodeURIComponent(channelContext.channelId)}/messages/${encodeURIComponent(threadMessageId)}/replies`;
    let request = this.graphClient.api(apiBase).top(limit).orderby("createdDateTime desc");
    if (cursor) {
      request = request.filter(`createdDateTime lt ${cursor}`);
    }
    const response = await request.get();
    const graphMessages = response.value ?? [];
    graphMessages.reverse();
    const messages = this.graphMessagesToMessages(graphMessages, threadId);
    let nextCursor;
    if (graphMessages.length >= limit) {
      const oldest = graphMessages[0];
      if (oldest?.createdDateTime) nextCursor = oldest.createdDateTime;
    }
    return { messages, nextCursor };
  }
  async fetchThread(threadId) {
    const { conversationId, serviceUrl } = this.decodeThreadId(threadId);
    return {
      id: threadId,
      channelId: this.channelIdFromThreadId(threadId),
      isDM: this.isDM(threadId),
      metadata: { conversationId, serviceUrl }
    };
  }
  async fetchChannelInfo(channelId) {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "fetchChannelInfo requires Microsoft Graph API access.",
        "fetchChannelInfo"
      );
    }
    const parts = channelId.split(":");
    const teamId = parts[0];
    const channelPart = parts.slice(1).join(":");
    const response = await this.graphClient.api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}`).get();
    return {
      id: channelId,
      name: response.displayName ?? channelId,
      metadata: {
        description: response.description,
        isPrivate: response.membershipType === "private",
        createdDateTime: response.createdDateTime
      }
    };
  }
  async fetchChannelMessages(channelId, options = {}) {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "fetchChannelMessages requires Microsoft Graph API access.",
        "fetchChannelMessages"
      );
    }
    const parts = channelId.split(":");
    const teamId = parts[0];
    const channelPart = parts.slice(1).join(":");
    const limit = options.limit ?? 50;
    const response = await this.graphClient.api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}/messages`).top(limit).orderby("createdDateTime desc").get();
    const graphMessages = response.value ?? [];
    graphMessages.reverse();
    const threadId = `teamssdk:${channelId}:channel`;
    const messages = this.graphMessagesToMessages(graphMessages, threadId);
    return { messages };
  }
  async postChannelMessage(channelId, message) {
    const parts = channelId.split(":");
    const teamId = parts[0];
    const channelPart = parts.slice(1).join(":");
    const card = extractCard(message);
    const threadId = `teamssdk:${channelId}:channel`;
    if (card) {
      const adaptiveCard = cardToAdaptiveCard(card);
      const payload2 = {
        body: { contentType: "html", content: '<attachment id="card1"></attachment>' },
        attachments: [
          {
            id: "card1",
            contentType: "application/vnd.microsoft.card.adaptive",
            content: JSON.stringify(adaptiveCard)
          }
        ]
      };
      if (!this.graphClient) {
        throw new NotImplementedError(
          "postChannelMessage with cards requires Graph API access",
          "postChannelMessage"
        );
      }
      const response2 = await this.graphClient.api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}/messages`).post(payload2);
      return { id: response2.id, threadId, raw: payload2 };
    }
    const text = convertEmojiPlaceholders(
      this.formatConverter.renderPostable(message),
      "teams"
    );
    if (!this.graphClient) {
      throw new NotImplementedError(
        "postChannelMessage requires Graph API access",
        "postChannelMessage"
      );
    }
    const payload = { body: { contentType: "html", content: text } };
    const response = await this.graphClient.api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}/messages`).post(payload);
    return { id: response.id, threadId, raw: payload };
  }
  async listThreads(channelId, _options) {
    if (!this.graphClient) {
      throw new NotImplementedError(
        "listThreads requires Microsoft Graph API access.",
        "listThreads"
      );
    }
    const parts = channelId.split(":");
    const teamId = parts[0];
    const channelPart = parts.slice(1).join(":");
    const response = await this.graphClient.api(`/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelPart)}/messages`).top(50).orderby("createdDateTime desc").get();
    const graphMessages = response.value ?? [];
    const threads = graphMessages.map((msg) => {
      const rootThreadId = this.encodeThreadId({
        conversationId: `${channelId};messageid=${msg.id}`,
        serviceUrl: ""
      });
      const rawText = msg.body?.content?.replace(/<[^>]+>/g, "") ?? "";
      const rootMessage = new Message({
        id: msg.id,
        threadId: rootThreadId,
        text: rawText,
        formatted: this.formatConverter.toAst(rawText),
        raw: msg,
        author: {
          userId: msg.from?.user?.id ?? msg.from?.application?.id ?? "unknown",
          userName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? "unknown",
          fullName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? "unknown",
          isBot: !!msg.from?.application,
          isMe: msg.from?.application?.id === this.config.appId
        },
        metadata: {
          dateSent: msg.createdDateTime ? new Date(msg.createdDateTime) : /* @__PURE__ */ new Date(),
          edited: !!msg.lastModifiedDateTime
        },
        attachments: []
      });
      return {
        id: rootThreadId,
        rootMessage,
        lastReplyAt: msg.lastModifiedDateTime ? new Date(msg.lastModifiedDateTime) : void 0
      };
    });
    return { threads };
  }
  parseMessage(raw) {
    const activity = raw;
    const threadId = this.encodeThreadId({
      conversationId: activity.conversation?.id ?? "",
      serviceUrl: activity.serviceUrl ?? ""
    });
    return this.parseTeamsMessage(activity, threadId);
  }
  renderFormatted(content) {
    return this.formatConverter.fromAst(content);
  }
  encodeThreadId(data) {
    const encoded = Buffer.from(JSON.stringify(data)).toString("base64url");
    return `teamssdk:${encoded}`;
  }
  decodeThreadId(threadId) {
    const prefix = "teamssdk:";
    const encoded = threadId.startsWith(prefix) ? threadId.slice(prefix.length) : threadId;
    try {
      return JSON.parse(
        Buffer.from(encoded, "base64url").toString("utf8")
      );
    } catch {
      return { conversationId: threadId, serviceUrl: "" };
    }
  }
  channelIdFromThreadId(threadId) {
    try {
      const { conversationId } = this.decodeThreadId(threadId);
      const base = conversationId.replace(MESSAGEID_STRIP_PATTERN, "");
      return base;
    } catch {
      return threadId;
    }
  }
  isDM(threadId) {
    try {
      const { conversationId } = this.decodeThreadId(threadId);
      return !conversationId.includes("@thread") && !conversationId.includes("@conference");
    } catch {
      return false;
    }
  }
  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------
  parseTeamsMessage(activity, threadId) {
    const text = activity.text ?? "";
    const normalizedText = text.trim();
    const isMe = this.isMessageFromSelf(activity);
    return new Message({
      id: activity.id ?? "",
      threadId,
      text: this.formatConverter.extractPlainText(normalizedText),
      formatted: this.formatConverter.toAst(normalizedText),
      raw: activity,
      author: {
        userId: activity.from?.id ?? "unknown",
        userName: activity.from?.name ?? "unknown",
        fullName: activity.from?.name ?? "unknown",
        isBot: activity.from?.role === "bot",
        isMe
      },
      metadata: {
        dateSent: activity.timestamp ? new Date(activity.timestamp) : /* @__PURE__ */ new Date(),
        edited: false
      },
      attachments: (activity.attachments ?? []).filter(
        (att) => att.contentType !== "application/vnd.microsoft.card.adaptive" && !(att.contentType === "text/html" && !att.contentUrl)
      ).map((att) => this.createAttachment(att))
    });
  }
  createAttachment(att) {
    const url = att.contentUrl;
    let type = "file";
    if (att.contentType?.startsWith("image/")) type = "image";
    else if (att.contentType?.startsWith("video/")) type = "video";
    else if (att.contentType?.startsWith("audio/")) type = "audio";
    return {
      type,
      url,
      name: att.name,
      mimeType: att.contentType,
      fetchData: url ? async () => {
        const response = await fetch(url);
        if (!response.ok) {
          throw new NetworkError(
            "teamssdk",
            `Failed to fetch file: ${response.status} ${response.statusText}`
          );
        }
        const arrayBuffer = await response.arrayBuffer();
        return Buffer.from(arrayBuffer);
      } : void 0
    };
  }
  graphMessagesToMessages(graphMessages, threadId) {
    return graphMessages.map((msg) => {
      const isFromBot = msg.from?.application?.id === this.config.appId || msg.from?.user?.id === this.config.appId;
      const rawText = this.extractTextFromGraphMessage(msg);
      return new Message({
        id: msg.id,
        threadId,
        text: this.formatConverter.extractPlainText(rawText),
        formatted: this.formatConverter.toAst(rawText),
        raw: msg,
        author: {
          userId: msg.from?.user?.id ?? msg.from?.application?.id ?? "unknown",
          userName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? "unknown",
          fullName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? "unknown",
          isBot: !!msg.from?.application,
          isMe: isFromBot
        },
        metadata: {
          dateSent: msg.createdDateTime ? new Date(msg.createdDateTime) : /* @__PURE__ */ new Date(),
          edited: !!msg.lastModifiedDateTime
        },
        attachments: this.extractAttachmentsFromGraphMessage(msg)
      });
    });
  }
  extractTextFromGraphMessage(msg) {
    const body = msg.body;
    if (!body) return "";
    if (body.contentType === "html") {
      return body.content?.replace(/<[^>]+>/g, "") ?? "";
    }
    return body.content ?? "";
  }
  extractAttachmentsFromGraphMessage(msg) {
    return (msg.attachments ?? []).filter(
      (att) => att.contentType !== "application/vnd.microsoft.card.adaptive"
    ).map((att) => ({
      type: "file",
      url: att.contentUrl,
      name: att.name,
      mimeType: att.contentType
    }));
  }
  isMessageFromSelf(activity) {
    return activity.from?.id === this.config.appId;
  }
  /** Map Teams API errors to typed adapter errors */
  handleTeamsError(error, _operation) {
    if (error instanceof Error) {
      const msg = error.message ?? "";
      if (msg.includes("429")) {
        const retryAfter = 10;
        throw new AdapterRateLimitError("teamssdk", retryAfter);
      }
      if (msg.includes("401")) {
        throw new AuthenticationError("teamssdk", msg);
      }
      if (msg.includes("403")) {
        throw new PermissionError("teamssdk", msg);
      }
    }
    throw error;
  }
};
function createTeamsSDKAdapter(config) {
  return new TeamsSDKAdapter(config);
}
export {
  TeamsApp,
  TeamsSDKAdapter,
  TeamsSDKFormatConverter,
  cardToAdaptiveCard,
  cardToFallbackText,
  createTeamsSDKAdapter
};
//# sourceMappingURL=index.js.map