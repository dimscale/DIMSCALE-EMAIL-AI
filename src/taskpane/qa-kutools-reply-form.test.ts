/* global console */
/* global globalThis */
import { afterEach, describe, expect, it, vi } from "vitest";
import { REPLY_FORM_LIMIT, toReplyHtml } from "./messageView";
import { openReplyForm } from "./replyForm";

// QA (independent): reply-form HTML built from AI text (F-003/AC-24, decision 5). Pure logic plus mocked Office.

const ENTITY = /^&(amp|lt|gt|quot|#39);/;

// Every raw character left in the output must be a <br> tag or part of a known entity.
const assertOnlySafeMarkup = (html: string) => {
  let i = 0;
  while (i < html.length) {
    const c = html[i];
    if (c === "<") {
      expect(html.startsWith("<br>", i)).toBe(true);
      i += 4;
    } else if (c === ">" || c === '"' || c === "'") {
      throw new Error(`raw ${c} at ${i}`);
    } else if (c === "&") {
      const m = ENTITY.exec(html.slice(i, i + 8));
      expect(m, `bare & at ${i}`).not.toBeNull();
      i += m![0].length;
    } else i++;
  }
};

// Deterministic PRNG so the fuzz is reproducible.
const rng = (seed: number) => () => {
  seed = (seed * 1664525 + 1013904223) >>> 0;
  return seed / 4294967296;
};

describe("toReplyHtml escaping (F-003/AC-24)", () => {
  it("escapes & < > \" ' exactly once and keeps an existing entity literal", () => {
    expect(toReplyHtml(`& < > " '`)).toBe("&amp; &lt; &gt; &quot; &#39;");
    expect(toReplyHtml("&amp;")).toBe("&amp;amp;");
    expect(toReplyHtml("&lt;script&gt;")).toBe("&amp;lt;script&amp;gt;");
  });

  it("turns each newline kind into one <br>: LF, CRLF, lone CR, and runs of them", () => {
    expect(toReplyHtml("a\nb")).toBe("a<br>b");
    expect(toReplyHtml("a\r\nb")).toBe("a<br>b");
    expect(toReplyHtml("a\rb")).toBe("a<br>b");
    expect(toReplyHtml("a\r\n\r\nb")).toBe("a<br><br>b");
    expect(toReplyHtml("a\n\rb")).toBe("a<br><br>b");
    expect(toReplyHtml("\n")).toBe("<br>");
  });

  it("keeps tabs, emoji, RTL and a lone surrogate as they are (no raw markup characters)", () => {
    expect(toReplyHtml("a\tb 📧 مرحبا")).toBe("a\tb 📧 مرحبا");
    expect(() => toReplyHtml("\uD800")).not.toThrow();
    expect(toReplyHtml("")).toBe("");
  });

  it("fuzz (seeded, 400 strings): no raw < > \" ' and no bare & in the output", () => {
    const alphabet = [
      "&",
      "<",
      ">",
      '"',
      "'",
      "\n",
      "\r",
      "a",
      " ",
      "😀",
      "\uD800",
      "\uDC00",
      "é",
      "<br>",
      "&amp;",
    ];
    const r = rng(20261029);
    for (let k = 0; k < 400; k++) {
      let s = "";
      const len = Math.floor(r() * 60);
      for (let j = 0; j < len; j++) s += alphabet[Math.floor(r() * alphabet.length)];
      const out = toReplyHtml(s);
      // a literal "<br>" typed by the model is escaped too, so the only raw "<" is one we added
      assertOnlySafeMarkup(out.replace(/<br>/g, "\u0001"));
      expect(out.length).toBeLessThanOrEqual(REPLY_FORM_LIMIT);
    }
  });
});

describe("toReplyHtml 32 KB boundary (F-003/AC-24, edge: 3,000 characters)", () => {
  it("exactly 32 KB of plain text is kept whole", () => {
    expect(toReplyHtml("a".repeat(REPLY_FORM_LIMIT))).toHaveLength(REPLY_FORM_LIMIT);
  });

  it("one character over is cut, never mid-entity", () => {
    const out = toReplyHtml("a".repeat(REPLY_FORM_LIMIT + 1));
    expect(out).toHaveLength(REPLY_FORM_LIMIT);
  });

  it("40,000 apostrophes: the cut falls on a whole &#39; (5 chars) and stays within the limit", () => {
    const out = toReplyHtml("'".repeat(40000));
    expect(out.length).toBeLessThanOrEqual(REPLY_FORM_LIMIT);
    expect(out.endsWith("&#39;")).toBe(true);
    expect(out.length % 5).toBe(0);
    expect(out.length).toBe(Math.floor(REPLY_FORM_LIMIT / 5) * 5); // 32765
  });

  it("newlines: the cut falls on a whole <br> (4 chars)", () => {
    const out = toReplyHtml("\n".repeat(20000));
    expect(out.length).toBeLessThanOrEqual(REPLY_FORM_LIMIT);
    expect(out.endsWith("<br>")).toBe(true);
    expect(out.length % 4).toBe(0);
  });

  it("the longest reply the function allows (3,000 characters, all quotes and ampersands) fits without truncation", () => {
    const reply = '"&'.repeat(1500); // 3,000 characters
    const out = toReplyHtml(reply);
    expect(out).toHaveLength(1500 * 6 + 1500 * 5);
    expect(out.length).toBeLessThan(REPLY_FORM_LIMIT);
    expect(out.startsWith("&quot;&amp;")).toBe(true);
  });

  it("limit is counted in UTF-16 units, not UTF-8 bytes (noted: 3,000-char replies cannot reach either limit)", () => {
    // 3,000 CJK characters are 9,000 UTF-8 bytes; the limit check is on string length, so the byte size is not checked
    const out = toReplyHtml("字".repeat(3000));
    expect(out).toHaveLength(3000);
  });
});

describe("openReplyForm error paths (F-003/AC-24), mocked Office", () => {
  afterEach(() => {
    delete (globalThis as any).Office;
  });

  it("returns false when displayReplyFormAsync is missing on a 1.9 host (no throw, nothing logged)", async () => {
    (globalThis as any).Office = {
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      context: { requirements: { isSetSupported: () => true } },
    };
    const spy = vi.spyOn(console, "error").mockImplementation(() => undefined);
    const logSpy = vi.spyOn(console, "log").mockImplementation(() => undefined);
    expect(await openReplyForm({} as any, "x")).toBe(false);
    expect(spy).not.toHaveBeenCalled();
    expect(logSpy).not.toHaveBeenCalled();
    spy.mockRestore();
    logSpy.mockRestore();
  });

  it("returns false when Office.context.requirements is missing (old host), without a throw", async () => {
    (globalThis as any).Office = {
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      context: {},
    };
    expect(await openReplyForm({ displayReplyForm: vi.fn() } as any, "x")).toBe(false);
  });

  it("passes the escaped html string unchanged to displayReplyFormAsync (1.9), once", async () => {
    const calls: string[] = [];
    (globalThis as any).Office = {
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      context: { requirements: { isSetSupported: () => true } },
    };
    const item = {
      displayReplyFormAsync: (h: string, cb: (r: { status: string }) => void) => {
        calls.push(h);
        cb({ status: "succeeded" });
      },
      displayReplyForm: vi.fn(),
    };
    expect(await openReplyForm(item as any, toReplyHtml("<x>\nk&"))).toBe(true);
    expect(calls).toEqual(["&lt;x&gt;<br>k&amp;"]);
    expect(item.displayReplyForm).not.toHaveBeenCalled();
  });
});
