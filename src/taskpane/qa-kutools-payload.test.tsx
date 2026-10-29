import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

// QA (independent): exact request payload per action (F-003/AC-18, AC-35, AC-37, AC-41). Mocks only.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";
import { ANALYZE_URL } from "./config";

type Cb = (r: { status: string; value?: string }) => void;
let office: any;
let fetchMock: ReturnType<typeof vi.fn>;

// A high surrogate not followed by a low one, or a low surrogate not preceded by a high one.
const LONE_HIGH = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])/;
const LONE_LOW = /(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/;

const makeItem = (body: string, o: object = {}) => ({
  subject: "Pedido",
  from: { displayName: "Ana Silva", emailAddress: "ANA@Exemplo.PT " },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  ...o,
});
const setup = (item: unknown, extra: object = {}) => {
  office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      roamingSettings: { get: vi.fn(() => undefined), set: vi.fn(), saveAsync: vi.fn() },
      mailbox: {
        item,
        // Identity and recipients the pane must never send
        userProfile: { emailAddress: "thiago.santos@dimscale.com", displayName: "Thiago Santos" },
        addHandlerAsync: vi.fn(),
        removeHandlerAsync: vi.fn(),
        ...extra,
      },
    },
  };
  (globalThis as any).Office = office;
};
const json = (status: number, body: unknown) => ({
  ok: status < 300,
  status,
  json: async () => body,
});
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name }));
  });
};
const tab = (name: string) => fireEvent.click(screen.getByRole("tab", { name }));
const sent = (i = 0) => JSON.parse(fetchMock.mock.calls[i][1].body);
const okSummary = () =>
  json(200, {
    ok: true,
    action: "summarize",
    summary: {
      points: ["a", "b", "c"],
      type: "request",
      priority: "high",
      tone: "neutral",
      deadline: null,
      requestedAction: null,
    },
  });

beforeEach(() => {
  fetchMock = vi.fn(async () => okSummary());
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("summarize payload (F-003/AC-18)", () => {
  it("exact key set and values; domain after the last @, trimmed, lower-cased; body trimmed; URL, method and headers exact", async () => {
    setup(makeItem("  \n corpo \t "));
    render(<App />);
    await click("Resumir");
    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(ANALYZE_URL);
    expect(init.method).toBe("POST");
    expect(init.headers).toEqual({
      Authorization: "Bearer TOK",
      "Content-Type": "application/json",
    });
    expect(sent()).toEqual({
      action: "summarize",
      subject: "Pedido",
      fromName: "Ana Silva",
      fromDomain: "exemplo.pt",
      body: "corpo",
    });
  });

  it("sender without display name gives fromName ''; address without @ gives fromDomain ''; missing from gives both ''", async () => {
    setup(makeItem("x", { from: { displayName: "", emailAddress: "sem-arroba" } }));
    render(<App />);
    await click("Resumir");
    expect(sent()).toMatchObject({ fromName: "", fromDomain: "" });
  });

  it("missing from object and missing subject do not throw and send empty strings", async () => {
    setup(makeItem("x", { from: undefined, subject: undefined }));
    render(<App />);
    await click("Resumir");
    expect(sent()).toMatchObject({ subject: "", fromName: "", fromDomain: "" });
  });

  it("only the part after the LAST @ is sent as domain (a@b@C.Org -> c.org)", async () => {
    setup(makeItem("x", { from: { displayName: "X", emailAddress: "a@b@C.Org" } }));
    render(<App />);
    await click("Resumir");
    expect(sent().fromDomain).toBe("c.org");
    expect(JSON.stringify(sent())).not.toContain("a@b");
  });

  it("body: 2000 and 2500 characters are sent in full (not the 2,000 preview); 4000 exact; above 4000 cut to 4000", async () => {
    for (const n of [2000, 2001, 2500, 4000, 4001, 9000]) {
      cleanup();
      fetchMock.mockClear();
      setup(makeItem("x".repeat(n)));
      render(<App />);
      await click("Resumir");
      expect(sent().body).toHaveLength(Math.min(n, 4000));
    }
  });

  it("a 4,000-limit cut never leaves a lone high surrogate: an emoji straddling 4000 is dropped, the JSON has no lone surrogate", async () => {
    setup(makeItem("a".repeat(3999) + "😀" + "tail"));
    render(<App />);
    await click("Resumir");
    const body = sent().body as string;
    expect(body).toHaveLength(3999);
    expect(body.endsWith("a")).toBe(true);
    expect(LONE_HIGH.test(body)).toBe(false);
    expect(LONE_LOW.test(body)).toBe(false);
    expect(LONE_HIGH.test(fetchMock.mock.calls[0][1].body)).toBe(false);
  });

  it("an emoji that fits entirely inside 4000 is kept whole", async () => {
    setup(makeItem("a".repeat(3998) + "😀"));
    render(<App />);
    await click("Resumir");
    expect(sent().body).toHaveLength(4000);
    expect(sent().body.endsWith("😀")).toBe(true);
  });

  it("subject straddling 200 and name straddling 100 are cut without a lone surrogate", async () => {
    setup(
      makeItem("x", {
        subject: "s".repeat(199) + "😀",
        from: { displayName: "n".repeat(99) + "😀", emailAddress: "a@b.pt" },
      })
    );
    render(<App />);
    await click("Resumir");
    expect(sent().subject).toHaveLength(199);
    expect(sent().fromName).toHaveLength(99);
    expect(LONE_HIGH.test(sent().subject)).toBe(false);
    expect(LONE_HIGH.test(sent().fromName)).toBe(false);
  });

  it("a subject of 200 and a name of 100 characters are sent unchanged", async () => {
    setup(
      makeItem("x", {
        subject: "s".repeat(200),
        from: { displayName: "n".repeat(100), emailAddress: "a@b.pt" },
      })
    );
    render(<App />);
    await click("Resumir");
    expect(sent().subject).toHaveLength(200);
    expect(sent().fromName).toHaveLength(100);
  });
});

describe("reply and translate payloads (F-003/AC-35, AC-41)", () => {
  it("reply: exact key set; defaults match + none; URL, method and headers exact", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "reply", reply: "r" }));
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(ANALYZE_URL);
    expect(init.method).toBe("POST");
    expect(init.headers).toEqual({
      Authorization: "Bearer TOK",
      "Content-Type": "application/json",
    });
    expect(sent()).toEqual({
      action: "reply",
      subject: "Pedido",
      fromName: "Ana Silva",
      fromDomain: "exemplo.pt",
      body: "corpo",
      language: "match",
      tone: "none",
    });
  });

  it("translate: exact key set, fromName and fromDomain are the empty string even when the sender exists, default targetLanguage auto", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(
      json(200, { ok: true, action: "translate", translation: { subject: "", body: "b" } })
    );
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    expect(sent()).toEqual({
      action: "translate",
      subject: "Pedido",
      fromName: "",
      fromDomain: "",
      body: "corpo",
      targetLanguage: "auto",
    });
  });

  it("translate target select offers exactly auto, pt, en and each value is sent as is", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(
      json(200, { ok: true, action: "translate", translation: { subject: "", body: "b" } })
    );
    render(<App />);
    tab("Traduzir");
    const sel = screen.getByLabelText("Traduzir para") as HTMLSelectElement;
    expect(Array.from(sel.options).map((o) => o.value)).toEqual(["auto", "pt", "en"]);
    for (const v of ["auto", "pt", "en"]) {
      fireEvent.change(sel, { target: { value: v } });
      await click("Traduzir");
    }
    expect(fetchMock.mock.calls.map((_, i) => sent(i).targetLanguage)).toEqual([
      "auto",
      "pt",
      "en",
    ]);
  });

  it("translate sends the 4,000 cut body with no lone surrogate and no sender data (AC-37)", async () => {
    setup(makeItem("b".repeat(3999) + "😀"), {});
    fetchMock.mockResolvedValue(
      json(200, { ok: true, action: "translate", translation: { subject: "", body: "b" } })
    );
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    expect(sent().body).toHaveLength(3999);
    expect(LONE_HIGH.test(sent().body)).toBe(false);
    const raw = fetchMock.mock.calls[0][1].body as string;
    expect(raw).not.toContain("Ana");
    expect(raw).not.toContain("exemplo");
    expect(raw).not.toContain("ANA@");
  });
});

describe("no identifiers, identity or recipients in any request body", () => {
  const KEYS: Record<string, string[]> = {
    summarize: ["action", "body", "fromDomain", "fromName", "subject"],
    reply: ["action", "body", "fromDomain", "fromName", "language", "subject", "tone"],
    translate: ["action", "body", "fromDomain", "fromName", "subject", "targetLanguage"],
  };
  it.each([
    ["summarize", "Resumir", null],
    ["reply", "Sugerir resposta", "Resposta"],
    ["translate", "Traduzir", "Traduzir"],
  ] as const)(
    "%s body has exactly its keys and no itemId, recipients, dates, user identity",
    async (action, button, tabName) => {
      setup(
        makeItem("corpo", {
          itemId: "AAMkADSECRET-ITEM-ID-123",
          to: [{ emailAddress: "destinatario@secreto.example" }],
          cc: [{ emailAddress: "copia@secreto.example" }],
          conversationId: "CONV-SECRET-999",
          internetMessageId: "<msgid-secret@example>",
        })
      );
      fetchMock.mockResolvedValue(
        action === "summarize"
          ? okSummary()
          : action === "reply"
            ? json(200, { ok: true, action: "reply", reply: "r" })
            : json(200, { ok: true, action: "translate", translation: { subject: "", body: "b" } })
      );
      render(<App />);
      if (tabName) tab(tabName);
      await click(button);
      const raw = fetchMock.mock.calls[0][1].body as string;
      for (const secret of [
        "AAMkADSECRET",
        "CONV-SECRET",
        "msgid-secret",
        "destinatario",
        "copia@",
        "secreto.example",
        "thiago.santos",
        "dimscale.com",
        "Thiago",
      ])
        expect(raw).not.toContain(secret);
      expect(Object.keys(JSON.parse(raw)).sort()).toEqual(KEYS[action].sort());
    }
  );
});
