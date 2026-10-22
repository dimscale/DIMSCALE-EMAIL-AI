/* global Office */

import { describe, expect, it } from "vitest";
import { BODY_LIMIT, toMessageView } from "./messageView";

const item = (o: object = {}) =>
  ({
    subject: "Hi",
    from: { displayName: "Ann", emailAddress: "ann@example.com" },
    dateTimeCreated: new Date(2026, 0, 2, 3, 4),
    ...o,
  }) as unknown as Office.MessageRead;

describe("toMessageView", () => {
  it("maps subject, sender and date (AC-3)", () => {
    const v = toMessageView(item(), "body");
    expect(v).toMatchObject({
      subject: "Hi",
      senderName: "Ann",
      senderEmail: "ann@example.com",
      body: "body",
    });
    expect(v.date).toBe(new Date(2026, 0, 2, 3, 4).toLocaleString());
  });
  it("truncates over 2000 chars and keeps full length (AC-4)", () => {
    const v = toMessageView(item(), "a".repeat(2001));
    expect(v.body).toHaveLength(2000);
    expect(v.truncated).toBe(true);
    expect(v.bodyLength).toBe(2001);
  });
  it("does not truncate at exactly 2000 (AC-4)", () => {
    const v = toMessageView(item(), "a".repeat(2000));
    expect(v.truncated).toBe(false);
    expect(v.bodyLength).toBe(2000);
  });
  it("defaults empty subject and handles empty body (AC-5)", () => {
    const v = toMessageView(item({ subject: "  " }), "");
    expect(v.subject).toBe("(no subject)");
    expect(v.body).toBe("");
  });
  it("handles a sender without display name", () => {
    const v = toMessageView(item({ from: { emailAddress: "x@example.com" } }), "b");
    expect(v.senderName).toBe("");
    expect(v.senderEmail).toBe("x@example.com");
  });
});

// Added in QA pass for F-001 (edge cases from the spec and the QA brief).
describe("toMessageView edge cases (QA)", () => {
  it("truncates a body of 1.5 MB to 2000 chars and keeps the full length (AC-4, large body)", () => {
    const v = toMessageView(item(), "x".repeat(1_500_000));
    expect(v.body).toHaveLength(BODY_LIMIT);
    expect(v.truncated).toBe(true);
    expect(v.bodyLength).toBe(1_500_000);
  });

  it("keeps emoji, accents and CJK in subject and body literally (edge: non-ASCII)", () => {
    const subject = "Olá 👋 — 你好 café";
    const body = "Olá 👋 ünïcödé 日本語 🎉";
    const v = toMessageView(item({ subject }), body);
    expect(v.subject).toBe(subject);
    expect(v.body).toBe(body);
    expect(v.bodyLength).toBe(body.length);
    expect(v.truncated).toBe(false);
  });

  it("does not leave a lone surrogate when the 2000-char cut splits an emoji (AC-4)", () => {
    const text = "a".repeat(BODY_LIMIT - 1) + "🎉" + "b";
    const v = toMessageView(item(), text);
    expect(v.body).not.toMatch(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/);
  });

  it("treats a null body as empty (defensive)", () => {
    const v = toMessageView(item(), null as unknown as string);
    expect(v.body).toBe("");
    expect(v.bodyLength).toBe(0);
    expect(v.truncated).toBe(false);
  });

  it("handles a sender with neither display name nor address without throwing (edge)", () => {
    const v = toMessageView(item({ from: { displayName: "", emailAddress: "" } }), "b");
    expect(v.senderName).toBe("");
    expect(v.senderEmail).toBe("");
  });

  it("handles a missing from object and missing date (edge)", () => {
    const v = toMessageView(item({ from: undefined, dateTimeCreated: undefined }), "b");
    expect(v.senderName).toBe("");
    expect(v.senderEmail).toBe("");
    expect(v.date).toBe("");
  });
});
