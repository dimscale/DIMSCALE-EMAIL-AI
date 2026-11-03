import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
vi.mock("./config", async (orig) => ({
  ...(await orig<Record<string, unknown>>()),
  AUTO_MODE_AVAILABLE: true,
}));
import App from "./App";
import { getTokenSilent } from "./auth";

// F-003 amendment (j): suggestion kept per email while the pane is open (memory only).
type Cb = (r: { status: string; value?: string }) => void;
const SUMMARY = {
  points: ["ponto <b>um</b>", "dois", "tres"],
  type: "request",
  priority: "high",
  tone: "neutral",
  deadline: null,
  requestedAction: null,
};
const handlers = new Set<() => void>(); // the theme hook registers its own handler too
let office: any;
let rs: { get: any; set: any; saveAsync: any };
let fetchMock: ReturnType<typeof vi.fn>;

const makeItem = (body: string, o: object = {}) => ({
  subject: "Assunto",
  from: { displayName: "Ana", emailAddress: "ana@exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  ...o,
});
const setup = (item: unknown, stored: unknown = true, saveStatus = "succeeded") => {
  handlers.clear();
  rs = {
    get: vi.fn((k: string) => (k === "autoSummary" ? stored : undefined)),
    set: vi.fn(),
    saveAsync: vi.fn((cb: Cb) => cb({ status: saveStatus })),
  };
  office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      roamingSettings: rs,
      mailbox: {
        item,
        addHandlerAsync: vi.fn((e: string, h: () => void) => {
          if (e === "itemChanged") handlers.add(h);
        }),
        removeHandlerAsync: vi.fn(),
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
const summaryOk = () => json(200, { ok: true, action: "summarize", summary: SUMMARY });
const advance = async (ms: number) => {
  await act(async () => {
    await vi.advanceTimersByTimeAsync(ms);
  });
};
const change = async (item: unknown) => {
  office.context.mailbox.item = item;
  await act(async () => handlers.forEach((h) => h()));
};
const tab = (name: string) => fireEvent.click(screen.getByRole("tab", { name }));
const sentBodies = () => fetchMock.mock.calls.map((c) => JSON.parse(c[1].body));

const REPLY = { ok: true, action: "reply", reply: "resposta um" };
const TRANS = { ok: true, action: "translate", translation: { subject: "S", body: "traduzido" } };
const item = (n: number) => makeItem(`corpo ${n}`, { itemId: `id${n}` });
const click = async (name: string) => {
  fireEvent.click(screen.getByRole("button", { name }));
  await advance(0);
};

beforeEach(() => {
  vi.useFakeTimers();
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
  localStorage.clear();
  sessionStorage.clear();
});
afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("kept results (F-003 amendment j)", () => {
  it("revisiting an email shows its summary at once with zero requests; a click makes exactly one and replaces it", async () => {
    setup(item(1), false);
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "summarize", summary: SUMMARY }));
    render(<App />);
    await click("Resumir");
    expect(fetchMock).toHaveBeenCalledTimes(1);
    await change(item(2));
    expect(screen.queryByText("dois")).toBeNull();
    await change(item(1));
    await advance(2000);
    expect(screen.getByText("dois")).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    fetchMock.mockResolvedValue(
      json(200, { ok: true, action: "summarize", summary: { ...SUMMARY, points: ["novo"] } })
    );
    await click("Resumir");
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(screen.getByText("novo")).toBeTruthy();
    await change(item(2));
    await change(item(1));
    expect(screen.getByText("novo")).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("reply: revisit shows it with no request and keeps the user's current Idioma; a refined version is the one kept", async () => {
    setup(item(1), false);
    render(<App />);
    tab("Resposta");
    fireEvent.change(screen.getByLabelText("Idioma"), { target: { value: "en" } });
    fetchMock.mockResolvedValue(json(200, REPLY));
    await click("Sugerir resposta");
    fetchMock.mockResolvedValue(json(200, { ...REPLY, reply: "versao dois" }));
    fireEvent.change(screen.getByPlaceholderText(/./), { target: { value: "mais curto" } });
    await click("Enviar alterações");
    expect(fetchMock).toHaveBeenCalledTimes(2);
    await change(item(2));
    fireEvent.change(screen.getByLabelText("Idioma"), { target: { value: "pt" } });
    expect(screen.queryByText("versao dois")).toBeNull();
    await change(item(1));
    expect(screen.getByText("versao dois")).toBeTruthy();
    expect((screen.getByLabelText("Idioma") as HTMLSelectElement).value).toBe("pt"); // the user's last choice is not overwritten
    expect(screen.queryByRole("button", { name: /versão anterior/i })).toBeNull(); // history reset
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("translation: revisit shows it with no request and keeps the user's current target", async () => {
    setup(item(1), false);
    render(<App />);
    tab("Traduzir");
    fireEvent.change(screen.getByLabelText("Traduzir para"), { target: { value: "en" } });
    fetchMock.mockResolvedValue(json(200, TRANS));
    await click("Traduzir");
    await change(item(2));
    fireEvent.change(screen.getByLabelText("Traduzir para"), { target: { value: "pt" } });
    await change(item(1));
    expect(screen.getByText("traduzido")).toBeTruthy();
    expect((screen.getByLabelText("Traduzir para") as HTMLSelectElement).value).toBe("pt");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("auto mode does not run for a pair that has a kept result", async () => {
    setup(item(1), false);
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "summarize", summary: SUMMARY }));
    render(<App />);
    await click("Resumir");
    cleanup();
    setup(item(1), true); // auto summary on
    render(<App />);
    await advance(2000);
    expect(fetchMock).toHaveBeenCalledTimes(2); // new pane session: auto may run once
    await change(item(2));
    await advance(2000);
    expect(fetchMock).toHaveBeenCalledTimes(3);
    await change(item(1));
    await advance(2000);
    expect(fetchMock).toHaveBeenCalledTimes(3); // kept for item 1: no auto request
    expect(screen.getByText("dois")).toBeTruthy();
  });

  it("keeps at most 50 emails, dropping the oldest", async () => {
    setup(item(1), false);
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "summarize", summary: SUMMARY }));
    render(<App />);
    await click("Resumir");
    for (let n = 2; n <= 51; n++) {
      await change(item(n));
      await click("Resumir");
    }
    expect(fetchMock).toHaveBeenCalledTimes(51);
    await change(item(51));
    expect(screen.getByText("dois")).toBeTruthy();
    await change(item(2));
    expect(screen.getByText("dois")).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(51);
    await change(item(1)); // evicted: nothing shown, nothing requested
    expect(screen.queryByText("dois")).toBeNull();
    expect(fetchMock).toHaveBeenCalledTimes(51);
  });

  it("writes nothing to storage and logs nothing", async () => {
    const spies = (["log", "info", "warn", "error", "debug"] as const).map((m) =>
      vi.spyOn(console, m).mockImplementation(() => undefined)
    );
    setup(item(1), false);
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "summarize", summary: SUMMARY }));
    render(<App />);
    await click("Resumir");
    await change(item(2));
    await change(item(1));
    expect(rs.set).not.toHaveBeenCalled();
    expect(rs.saveAsync).not.toHaveBeenCalled();
    expect(localStorage.length).toBe(0);
    expect(sessionStorage.length).toBe(0);
    spies.forEach((s) => expect(s).not.toHaveBeenCalled());
    vi.restoreAllMocks();
  });
});
