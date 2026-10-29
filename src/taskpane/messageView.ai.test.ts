import { describe, expect, it } from "vitest";
import { fromDomain, toMessageView, toReplyHtml } from "./messageView";

const item = (o: object = {}) =>
  ({
    subject: "S",
    from: { displayName: "Ann", emailAddress: "Ana@Exemplo.PT " },
    dateTimeCreated: new Date(),
    ...o,
  }) as any;

describe("AI payload helpers (F-003/AC-18)", () => {
  it("derives the domain after the last @, lower-cased, empty without @", () => {
    expect(fromDomain("Ana@Exemplo.PT ")).toBe("exemplo.pt");
    expect(fromDomain("a@b@C.com")).toBe("c.com");
    expect(fromDomain("/o=x/cn=ana")).toBe("");
    expect(fromDomain(undefined)).toBe("");
  });
  it("cuts aiText to 4000 from the full body, trimmed, without a lone surrogate", () => {
    const v = toMessageView(item(), "  " + "a".repeat(3999) + "😀tail");
    expect(v.aiText).toHaveLength(3999);
    expect(v.aiText).not.toMatch(/[\uD800-\uDBFF]$/);
    expect(toMessageView(item(), "b".repeat(9000)).aiText).toHaveLength(4000);
    expect(toMessageView(item(), "  \n ").aiText).toBe("");
  });
  it("cuts subject (200) and name (100), keeps an empty subject empty", () => {
    const v = toMessageView(
      item({
        subject: "x".repeat(300),
        from: { displayName: "n".repeat(150), emailAddress: "a@b.pt" },
      }),
      "b"
    );
    expect(v.aiSubject).toHaveLength(200);
    expect(v.aiFromName).toHaveLength(100);
    expect(toMessageView(item({ subject: "" }), "b").aiSubject).toBe("");
  });
});

describe("toReplyHtml (F-003/AC-24)", () => {
  it("escapes & < > \" ' and turns every newline kind into <br>", () => {
    expect(toReplyHtml(`a & b <i> "q" 'r'\r\nx\ny\rz`)).toBe(
      "a &amp; b &lt;i&gt; &quot;q&quot; &#39;r&#39;<br>x<br>y<br>z"
    );
  });
  it("stays within 32 KB even when fully escaped", () => {
    expect(toReplyHtml("&".repeat(20000)).length).toBeLessThanOrEqual(32 * 1024);
  });
});
