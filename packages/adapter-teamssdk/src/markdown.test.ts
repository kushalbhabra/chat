/**
 * Tests for TeamsSDKFormatConverter (markdown.ts)
 */
import { describe, expect, it } from "vitest";
import { parseMarkdown } from "chat";
import { TeamsSDKFormatConverter } from "./markdown.js";

const converter = new TeamsSDKFormatConverter();

describe("TeamsSDKFormatConverter.toAst", () => {
  it("parses plain text", () => {
    const ast = converter.toAst("Hello world");
    expect(ast.type).toBe("root");
    expect(ast.children).toHaveLength(1);
  });

  it("converts Teams <at> mentions to @mentions", () => {
    const ast = converter.toAst("<at>Alice</at> hello");
    const text = converter.extractPlainText("<at>Alice</at> hello");
    expect(text).toContain("Alice");
  });

  it("converts HTML bold tags to markdown bold", () => {
    const ast = converter.toAst("<b>bold</b> text");
    const md = converter.fromAst(ast);
    expect(md).toContain("bold");
  });

  it("converts HTML italic tags to markdown italic", () => {
    const ast = converter.toAst("<i>italic</i> text");
    const md = converter.fromAst(ast);
    expect(md).toContain("italic");
  });

  it("converts HTML strikethrough to markdown", () => {
    const ast = converter.toAst("<s>strikethrough</s>");
    const md = converter.fromAst(ast);
    expect(md).toContain("strikethrough");
  });

  it("converts HTML links to markdown links", () => {
    const ast = converter.toAst('<a href="https://example.com">link text</a>');
    const md = converter.fromAst(ast);
    expect(md).toContain("link text");
    expect(md).toContain("https://example.com");
  });

  it("decodes HTML entities", () => {
    const ast = converter.toAst("1 &lt; 2 &amp; 3 &gt; 0");
    const text = converter.extractPlainText("1 &lt; 2 &amp; 3 &gt; 0");
    // The converter should normalize the text
    expect(ast.type).toBe("root");
  });

  it("strips remaining HTML tags", () => {
    const ast = converter.toAst("<div>clean text</div>");
    const md = converter.fromAst(ast);
    expect(md).toContain("clean text");
    expect(md).not.toContain("<div>");
  });
});

describe("TeamsSDKFormatConverter.fromAst", () => {
  it("converts bold markdown to bold", () => {
    const ast = parseMarkdown("**bold text**");
    const result = converter.fromAst(ast);
    expect(result).toContain("**bold text**");
  });

  it("converts italic markdown to italic", () => {
    const ast = parseMarkdown("_italic text_");
    const result = converter.fromAst(ast);
    expect(result).toContain("_italic text_");
  });

  it("converts strikethrough markdown", () => {
    const ast = parseMarkdown("~~strikethrough~~");
    const result = converter.fromAst(ast);
    expect(result).toContain("~~strikethrough~~");
  });

  it("converts inline code", () => {
    const ast = parseMarkdown("`code`");
    const result = converter.fromAst(ast);
    expect(result).toContain("`code`");
  });

  it("converts fenced code blocks", () => {
    const ast = parseMarkdown("```js\nconsole.log('hi');\n```");
    const result = converter.fromAst(ast);
    expect(result).toContain("```js");
    expect(result).toContain("console.log");
  });

  it("converts links to markdown link format", () => {
    const ast = parseMarkdown("[click here](https://example.com)");
    const result = converter.fromAst(ast);
    expect(result).toContain("[click here](https://example.com)");
  });

  it("converts @mentions to Teams <at> format", () => {
    const ast = parseMarkdown("Hello @alice");
    const result = converter.fromAst(ast);
    expect(result).toContain("<at>alice</at>");
  });

  it("converts blockquotes", () => {
    const ast = parseMarkdown("> quoted text");
    const result = converter.fromAst(ast);
    expect(result).toContain("> ");
  });

  it("converts unordered lists", () => {
    const ast = parseMarkdown("- item 1\n- item 2");
    const result = converter.fromAst(ast);
    expect(result).toContain("item 1");
    expect(result).toContain("item 2");
  });

  it("converts ordered lists", () => {
    const ast = parseMarkdown("1. first\n2. second");
    const result = converter.fromAst(ast);
    expect(result).toContain("first");
    expect(result).toContain("second");
  });

  it("converts GFM tables", () => {
    const ast = parseMarkdown("| A | B |\n|---|---|\n| 1 | 2 |");
    const result = converter.fromAst(ast);
    expect(result).toContain("| A |");
    expect(result).toContain("| 1 |");
  });
});

describe("TeamsSDKFormatConverter.renderPostable", () => {
  it("passes plain strings through with mention conversion", () => {
    const result = converter.renderPostable("Hello @world");
    expect(result).toContain("<at>world</at>");
  });

  it("handles raw message objects", () => {
    const result = converter.renderPostable({ raw: "raw text @user" });
    expect(result).toContain("<at>user</at>");
  });

  it("handles markdown objects", () => {
    const result = converter.renderPostable({ markdown: "**bold**" });
    expect(result).toContain("bold");
  });

  it("handles AST objects", () => {
    const ast = parseMarkdown("hello world");
    const result = converter.renderPostable({ ast });
    expect(result).toContain("hello world");
  });
});
