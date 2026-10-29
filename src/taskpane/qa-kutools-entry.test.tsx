import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

// QA (independent): resumo.html entry (F-012/AC-50). The page is the same bundle; only location.pathname decides. Mocks only.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
let handlers: Set<() => void>;
let office: any;
let fetchMock: ReturnType<typeof vi.fn>;

const makeItem = (body: string) => ({
  subject: "Assunto",
  from: { displayName: "Ana", emailAddress: "ana@exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
});
const setup = (item: unknown) => {
  handlers = new Set();
  office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      roamingSettings: { get: vi.fn(() => undefined), set: vi.fn(), saveAsync: vi.fn() },
      mailbox: {
        item,
        addHandlerAsync: vi.fn((e: string, h: () => void) => {
          if (e === "itemChanged") handlers.add(h);
        }),
        removeHandlerAsync: vi.fn(),
      },
    },
    onReady: vi.fn(),
  };
  (globalThis as any).Office = office;
};
const json = (status: number, body: unknown) => ({
  ok: status < 300,
  status,
  json: async () => body,
});
const okSummary = () =>
  json(200, {
    ok: true,
    action: "summarize",
    summary: {
      points: ["um", "dois", "tres"],
      type: "request",
      priority: "high",
      tone: "neutral",
      deadline: null,
      requestedAction: null,
    },
  });
const flush = async () => {
  await act(async () => undefined);
  await act(async () => undefined);
};

beforeEach(() => {
  fetchMock = vi.fn(async () => okSummary());
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  window.history.pushState({}, "", "/");
});

describe("entry page state (F-012/AC-50)", () => {
  it("opens on Resumo, sends exactly one summarize request, and shows the privacy notice and assurance before anything else", async () => {
    setup(makeItem("corpo"));
    render(<App summaryEntry />);
    await flush();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(JSON.parse(fetchMock.mock.calls[0][1].body).action).toBe("summarize");
    expect(screen.getByRole("tab", { name: "Resumo" }).getAttribute("aria-selected")).toBe("true");
    expect(screen.getByText(/são enviados à OpenAI/)).toBeTruthy();
    expect(screen.getByText("Sem envio automático. O envio é sempre feito por si.")).toBeTruthy();
    expect(screen.getByText("Gerado por IA, pode estar errado.")).toBeTruthy();
  });

  it("an empty body on the entry page sends nothing and shows the no-text notice", async () => {
    setup(makeItem("   "));
    render(<App summaryEntry />);
    await flush();
    expect(fetchMock).not.toHaveBeenCalled();
    expect(
      screen.getAllByText("Esta mensagem não tem texto para analisar.").length
    ).toBeGreaterThan(0);
  });

  it("a null item on the entry page sends nothing", async () => {
    setup(null);
    render(<App summaryEntry />);
    await flush();
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("a later item change on the entry page never starts a summary", async () => {
    setup(makeItem("A"));
    render(<App summaryEntry />);
    await flush();
    office.context.mailbox.item = makeItem("B");
    await act(async () => handlers.forEach((h) => h()));
    await flush();
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("entry page: the summary request is the byte-identical click request (same keys, no page or route data)", async () => {
    setup(makeItem("corpo"));
    render(<App summaryEntry />);
    await flush();
    const raw = fetchMock.mock.calls[0][1].body as string;
    expect(Object.keys(JSON.parse(raw)).sort()).toEqual([
      "action",
      "body",
      "fromDomain",
      "fromName",
      "subject",
    ]);
    expect(raw).not.toContain("resumo.html");
    expect(raw).not.toContain("taskpane");
  });
});

describe("entry decision from the path only (F-012/AC-50, index.tsx)", () => {
  // Each case loads the real index.tsx with a stubbed location and counts the requests the entry causes.
  const cases: [string, number][] = [
    ["/resumo.html", 1],
    ["/resumo.html?x=1#y", 1],
    ["/addin/resumo.html", 1],
    ["/taskpane.html?resumo=1#resumo", 0],
    ["/taskpane.html#/resumo.html", 0],
    ["/myresumo.html", 0],
    ["/resumo.htmlx", 0],
  ];
  it.each(cases)(
    "pathname %s → %i request(s)",
    async (path, expected) => {
      window.history.pushState({}, "", path);
      document.body.innerHTML = '<div id="container"></div>';
      setup(makeItem("corpo"));
      let ready: (() => void) | undefined;
      office.onReady = vi.fn((cb: () => void) => (ready = cb));
      vi.resetModules();
      await import("./index");
      await act(async () => ready?.());
      await flush();
      expect(fetchMock).toHaveBeenCalledTimes(expected);
      document.body.innerHTML = "";
    },
    60000
  );
});
