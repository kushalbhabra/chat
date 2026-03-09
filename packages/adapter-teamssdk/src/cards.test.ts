/**
 * Tests for Teams SDK Adapter - cards.ts
 */
import {
  Actions,
  Button,
  Card,
  CardLink,
  CardText,
  Divider,
  Field,
  Fields,
  Image,
  LinkButton,
  Section,
} from "chat";
import { describe, expect, it } from "vitest";
import { cardToAdaptiveCard, cardToFallbackText } from "./cards.js";

describe("cardToAdaptiveCard", () => {
  it("creates a valid adaptive card structure", () => {
    const card = Card({ title: "Test" });
    const ac = cardToAdaptiveCard(card);
    expect(ac.type).toBe("AdaptiveCard");
    expect(ac.$schema).toBe("http://adaptivecards.io/schemas/adaptive-card.json");
    expect(ac.version).toBe("1.4");
    expect(ac.body).toBeInstanceOf(Array);
  });

  it("converts a card with title", () => {
    const card = Card({ title: "Welcome Message" });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(1);
    expect(ac.body[0]).toEqual({
      type: "TextBlock",
      text: "Welcome Message",
      weight: "bolder",
      size: "large",
      wrap: true,
    });
  });

  it("converts a card with title and subtitle", () => {
    const card = Card({ title: "Order Update", subtitle: "Your package is on its way" });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(2);
    expect(ac.body[1]).toEqual({
      type: "TextBlock",
      text: "Your package is on its way",
      isSubtle: true,
      wrap: true,
    });
  });

  it("converts a card with header image", () => {
    const card = Card({ title: "Product", imageUrl: "https://example.com/product.png" });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(2);
    expect(ac.body[1]).toEqual({
      type: "Image",
      url: "https://example.com/product.png",
      size: "stretch",
    });
  });

  it("converts text elements with styles", () => {
    const card = Card({
      children: [
        CardText("Regular text"),
        CardText("Bold text", { style: "bold" }),
        CardText("Muted text", { style: "muted" }),
      ],
    });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(3);
    expect(ac.body[0]).toEqual({ type: "TextBlock", text: "Regular text", wrap: true });
    expect(ac.body[1]).toEqual({ type: "TextBlock", text: "Bold text", wrap: true, weight: "bolder" });
    expect(ac.body[2]).toEqual({ type: "TextBlock", text: "Muted text", wrap: true, isSubtle: true });
  });

  it("converts image elements", () => {
    const card = Card({
      children: [Image({ url: "https://example.com/img.png", alt: "My image" })],
    });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(1);
    expect(ac.body[0]).toEqual({
      type: "Image",
      url: "https://example.com/img.png",
      altText: "My image",
      size: "auto",
    });
  });

  it("converts divider elements", () => {
    const card = Card({ children: [Divider()] });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(1);
    expect(ac.body[0]).toEqual({ type: "Container", separator: true, items: [] });
  });

  it("converts actions with buttons to card-level actions", () => {
    const card = Card({
      children: [
        Actions([
          Button({ id: "approve", label: "Approve", style: "primary" }),
          Button({ id: "reject", label: "Reject", style: "danger", value: "data-123" }),
          Button({ id: "skip", label: "Skip" }),
        ]),
      ],
    });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(0);
    expect(ac.actions).toHaveLength(3);
    expect(ac.actions![0]).toEqual({
      type: "Action.Submit",
      title: "Approve",
      data: { actionId: "approve", value: undefined },
      style: "positive",
    });
    expect(ac.actions![1]).toEqual({
      type: "Action.Submit",
      title: "Reject",
      data: { actionId: "reject", value: "data-123" },
      style: "destructive",
    });
    expect(ac.actions![2]).toEqual({
      type: "Action.Submit",
      title: "Skip",
      data: { actionId: "skip", value: undefined },
    });
  });

  it("converts link buttons to Action.OpenUrl", () => {
    const card = Card({
      children: [
        Actions([LinkButton({ url: "https://example.com/docs", label: "View Docs", style: "primary" })]),
      ],
    });
    const ac = cardToAdaptiveCard(card);
    expect(ac.actions).toHaveLength(1);
    expect(ac.actions![0]).toEqual({
      type: "Action.OpenUrl",
      title: "View Docs",
      url: "https://example.com/docs",
      style: "positive",
    });
  });

  it("converts fields to FactSet", () => {
    const card = Card({
      children: [
        Fields([
          Field({ label: "Status", value: "Active" }),
          Field({ label: "Priority", value: "High" }),
        ]),
      ],
    });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(1);
    expect(ac.body[0]).toEqual({
      type: "FactSet",
      facts: [
        { title: "Status", value: "Active" },
        { title: "Priority", value: "High" },
      ],
    });
  });

  it("wraps section children in a Container", () => {
    const card = Card({ children: [Section([CardText("Inside section")])] });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(1);
    expect(ac.body[0].type).toBe("Container");
    expect((ac.body[0].items as unknown[]).length).toBeGreaterThan(0);
  });

  it("converts CardLink to a TextBlock with markdown link", () => {
    const card = Card({
      children: [CardLink({ url: "https://example.com", label: "Click here" })],
    });
    const ac = cardToAdaptiveCard(card);
    expect(ac.body).toHaveLength(1);
    expect(ac.body[0]).toEqual({
      type: "TextBlock",
      text: "[Click here](https://example.com)",
      wrap: true,
    });
  });

  it("produces no actions property when no buttons present", () => {
    const card = Card({ title: "T" });
    const ac = cardToAdaptiveCard(card);
    expect(ac.actions).toBeUndefined();
  });
});

describe("cardToFallbackText", () => {
  it("generates fallback text with title", () => {
    const card = Card({
      title: "Order Update",
      subtitle: "Status changed",
      children: [
        CardText("Your order is ready"),
        Fields([
          Field({ label: "Order ID", value: "#1234" }),
          Field({ label: "Status", value: "Ready" }),
        ]),
      ],
    });
    const text = cardToFallbackText(card);
    expect(text).toContain("**Order Update**");
    expect(text).toContain("Status changed");
    expect(text).toContain("Your order is ready");
    expect(text).toContain("Order ID: #1234");
    expect(text).toContain("Status: Ready");
  });

  it("handles card with only title", () => {
    const card = Card({ title: "Simple Card" });
    const text = cardToFallbackText(card);
    expect(text).toBe("**Simple Card**");
  });
});
