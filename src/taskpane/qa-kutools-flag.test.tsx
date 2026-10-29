import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { readFileSync } from "node:fs";
import { join } from "node:path";

// QA (independent): F-003 AC-42 with the REAL config (no mock of ./config). Mocks only: Office, fetch, auth.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { AUTO_SUMMARY_AVAILABLE } from "./config";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
let handlers: Set<() => void>;
let rs: { get: any; set: any; saveAsync: any };
let office: any;
let fetchMock: ReturnType<typeof vi.fn>;

const makeItem = (body: string, o: object = {}) => ({
  subject: "Assunto",
  from: { displayName: "Ana", emailAddress: "ana@exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  ...o,
});
const setup = (item: unknown) => {
  handlers = new Set();
  // roamingSettings says ON. With the flag false this must have no effect at all.
  rs = {
    get: vi.fn(() => true),
    set: vi.fn(),
    saveAsync: vi.fn((cb: Cb) => cb({ status: "succeeded" })),
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
const analyzeCalls = () =>
  fetchMock.mock.calls.filter((c) => String(c[0]).includes("analyze-message"));

beforeEach(() => {
  fetchMock = vi.fn(async () => json(200, { ok: true, action: "summarize", summary: {} }));
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.clearAllMocks();
});

describe("auto mode flag off, real config (F-003/AC-42, AC-17)", () => {
  it("config.ts exports the literal false", () => {
    expect(AUTO_SUMMARY_AVAILABLE).toBe(false);
    expect(readFileSync(join(process.cwd(), "src/taskpane/config.ts"), "utf8")).toMatch(
      /export const AUTO_SUMMARY_AVAILABLE = false;/
    );
  });

  it("roamingSettings autoSummary=true: zero requests, zero token calls, no settings read, no switch, across item changes and tab switches for 2 minutes", async () => {
    vi.useFakeTimers();
    setup(makeItem("A"));
    render(<App />);
    await advance(1000);
    tab("Resposta");
    tab("Traduzir");
    tab("Resumo");
    for (const b of ["B", "C", "D"]) {
      await change(makeItem(b));
      await advance(1000);
    }
    await change(null);
    await advance(120000);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(getTokenSilent).not.toHaveBeenCalled();
    expect(rs.get).not.toHaveBeenCalled();
    expect(rs.set).not.toHaveBeenCalled();
    expect(rs.saveAsync).not.toHaveBeenCalled();
    expect(screen.queryByRole("switch")).toBeNull();
    expect(screen.queryByText(/Resumo automático/)).toBeNull();
    expect(screen.queryByText(/ligado/)).toBeNull();
  });

  it("Automático tab with the flag off: no analyze-message request (settings calls are not AI)", async () => {
    vi.useFakeTimers();
    setup(makeItem("A"));
    render(<App />);
    tab("Automático");
    tab("Resumo");
    await change(makeItem("B"));
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(0);
    expect(rs.get).not.toHaveBeenCalled();
  });

  it("entry page with the flag off and stored true: exactly one request (the entry counts as a click), none after item changes", async () => {
    vi.useFakeTimers();
    setup(makeItem("A"));
    render(<App summaryEntry />);
    await advance(0);
    expect(analyzeCalls()).toHaveLength(1);
    expect(JSON.parse(analyzeCalls()[0][1].body).action).toBe("summarize");
    await change(makeItem("B"));
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(1);
    expect(rs.get).not.toHaveBeenCalled();
  });
});
