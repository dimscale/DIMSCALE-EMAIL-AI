import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
vi.mock("./config", async (orig) => ({
  ...(await orig<Record<string, unknown>>()),
  AUTO_SUMMARY_AVAILABLE: true,
}));
import App from "./App";
import { getTokenSilent } from "./auth";

// F-003 AC-42 to AC-45 with the availability flag forced to true. Fake timers, mocked Office and fetch.
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
    get: vi.fn(() => stored),
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

/* A fetch that stays pending until resolved, and rejects on abort like the real one. */
const hangingFetch = () => {
  const calls: { signal: AbortSignal; resolve: (v: unknown) => void }[] = [];
  fetchMock.mockImplementation(
    (_u: string, init: { signal: AbortSignal }) =>
      new Promise((resolve, reject) => {
        calls.push({ signal: init.signal, resolve });
        init.signal.addEventListener("abort", () => reject(new Error("aborted")));
      })
  );
  return calls;
};

beforeEach(() => {
  vi.useFakeTimers();
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("switch and consent (F-003/AC-42, AC-45)", () => {
  it("is shown, on only for a stored boolean true, with the ligado line", () => {
    setup(makeItem("corpo"), true);
    render(<App />);
    expect((screen.getByRole("switch") as HTMLInputElement).checked).toBe(true);
    expect(screen.getByText(/Resumo automático ligado/)).toBeTruthy();
    expect(rs.get).toHaveBeenCalledWith("autoSummary");
    cleanup();
    for (const v of ["true", 1, null, false]) {
      setup(makeItem("corpo"), v);
      render(<App />);
      expect((screen.getByRole("switch") as HTMLInputElement).checked).toBe(false);
      cleanup();
    }
  });
  it("switching on opens the dialog; Cancelar leaves it off with no set and no request", async () => {
    setup(makeItem("corpo"), false);
    render(<App />);
    fireEvent.click(screen.getByRole("switch"));
    await advance(0);
    expect(screen.getByText("Ligar o resumo automático?")).toBeTruthy();
    expect(screen.getByText(/limite diário de 50 pedidos/)).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "Cancelar" }));
    await advance(2000);
    expect((screen.getByRole("switch") as HTMLInputElement).checked).toBe(false);
    expect(rs.set).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("Escape closes the dialog without turning it on", async () => {
    setup(makeItem("corpo"), false);
    render(<App />);
    fireEvent.click(screen.getByRole("switch"));
    await advance(0);
    fireEvent.keyDown(screen.getByRole("dialog"), { key: "Escape" });
    await advance(2000);
    expect(rs.set).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("Ligar saves once and starts auto mode; the dialog shows again after switching off and on", async () => {
    fetchMock.mockResolvedValue(summaryOk());
    setup(makeItem("corpo"), false);
    render(<App />);
    fireEvent.click(screen.getByRole("switch"));
    await advance(0);
    expect(screen.getByRole("dialog")).toBeTruthy();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    expect(rs.set).toHaveBeenCalledWith("autoSummary", true);
    expect(rs.saveAsync).toHaveBeenCalledTimes(1);
    expect((screen.getByRole("switch") as HTMLInputElement).checked).toBe(true);
    await advance(800);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    await act(async () => {
      fireEvent.click(screen.getByRole("switch")); // off: no dialog
    });
    expect(screen.queryByRole("dialog")).toBeNull();
    expect(rs.set).toHaveBeenLastCalledWith("autoSummary", false);
    expect(rs.saveAsync).toHaveBeenCalledTimes(2);
    fireEvent.click(screen.getByRole("switch")); // on again: dialog again
    await advance(0);
    expect(screen.getByText("Ligar o resumo automático?")).toBeTruthy();
  });
  it("a failed save keeps the choice for the session and shows the notice", async () => {
    setup(makeItem("corpo"), false, "failed");
    render(<App />);
    fireEvent.click(screen.getByRole("switch"));
    await advance(0);
    expect(screen.getByRole("dialog")).toBeTruthy();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    expect((screen.getByRole("switch") as HTMLInputElement).checked).toBe(true);
    expect(screen.getByText(/Não foi possível guardar a preferência/)).toBeTruthy();
  });
  it("a throwing saveAsync is non-fatal", async () => {
    setup(makeItem("corpo"), false);
    rs.saveAsync.mockImplementation(() => {
      throw new Error("x");
    });
    render(<App />);
    fireEvent.click(screen.getByRole("switch"));
    await advance(0);
    expect(screen.getByRole("dialog")).toBeTruthy();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    expect(screen.getByText(/Não foi possível guardar a preferência/)).toBeTruthy();
  });
  it("switching off cancels the pending debounce", async () => {
    setup(makeItem("corpo"), true);
    render(<App />);
    await advance(400);
    await act(async () => {
      fireEvent.click(screen.getByRole("switch"));
    });
    await advance(2000);
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("switching off aborts an in-flight auto request and discards its result", async () => {
    setup(makeItem("corpo"), true);
    const calls = hangingFetch();
    render(<App />);
    await advance(800);
    expect(calls).toHaveLength(1);
    await act(async () => {
      fireEvent.click(screen.getByRole("switch"));
    });
    expect(calls[0].signal.aborted).toBe(true);
    calls[0].resolve(summaryOk());
    await advance(10);
    expect(screen.queryByText("dois")).toBeNull();
    expect((screen.getByRole("button", { name: "Resumir" }) as HTMLButtonElement).disabled).toBe(
      false
    );
  });
});

describe("when it runs (F-003/AC-43)", () => {
  it("(a) one request 800 ms after load, identical to a click request", async () => {
    setup(makeItem("corpo"), true);
    fetchMock.mockResolvedValue(summaryOk());
    render(<App />);
    await advance(799);
    expect(fetchMock).not.toHaveBeenCalled();
    await advance(1);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    const auto = fetchMock.mock.calls[0][1].body;
    expect(screen.getByText("ponto <b>um</b>")).toBeTruthy();
    await advance(5000);
    expect(fetchMock).toHaveBeenCalledTimes(1); // once per item
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Resumir" }));
    });
    expect(fetchMock.mock.calls[1][1].body).toBe(auto);
    expect(JSON.parse(auto)).toEqual({
      action: "summarize",
      subject: "Assunto",
      fromName: "Ana",
      fromDomain: "exemplo.pt",
      body: "corpo",
    });
  });
  it("(b) ItemChanged to B sends one request for B after 800 ms; (c) A, B, C within 800 ms sends one for C", async () => {
    setup(makeItem("A"), true);
    fetchMock.mockResolvedValue(summaryOk());
    render(<App />);
    await advance(800);
    await change(makeItem("B"));
    await advance(300);
    await change(makeItem("C"));
    await advance(300);
    await change(makeItem("D"));
    await advance(799);
    expect(sentBodies().map((b) => b.body)).toEqual(["A"]);
    await advance(1);
    expect(sentBodies().map((b) => b.body)).toEqual(["A", "D"]);
  });
  it("(d) a request in flight for B is aborted when C arrives and its late result is never shown", async () => {
    setup(makeItem("B"), true);
    const calls = hangingFetch();
    render(<App />);
    await advance(800);
    await change(makeItem("C"));
    expect(calls[0].signal.aborted).toBe(true);
    calls[0].resolve(summaryOk());
    await advance(10);
    expect(screen.queryByText("dois")).toBeNull();
    await advance(800);
    expect(calls).toHaveLength(2);
    expect(calls[1].signal.aborted).toBe(false);
  });
  it("one request per email per pane session (A, B, back to A)", async () => {
    setup(makeItem("A", { itemId: "id-a" }), true);
    fetchMock.mockResolvedValue(summaryOk());
    render(<App />);
    await advance(800);
    await change(makeItem("B", { itemId: "id-b" }));
    await advance(800);
    await change(makeItem("A", { itemId: "id-a" }));
    await advance(2000);
    expect(sentBodies().map((b) => b.body)).toEqual(["A", "B"]);
  });
  it("(e) switching to Resumo sends one request when the item has no summary, none after", async () => {
    setup(makeItem("corpo"), true);
    fetchMock.mockResolvedValue(summaryOk());
    render(<App />);
    tab("Resposta"); // before the debounce ends: the timer is cancelled
    await advance(2000);
    expect(fetchMock).not.toHaveBeenCalled();
    tab("Resumo");
    await advance(800);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    tab("Resposta");
    tab("Resumo");
    await advance(2000);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it("(f) nothing on Resposta, Traduzir or Automático, even across item changes", async () => {
    setup(makeItem("A"), true);
    render(<App />);
    tab("Traduzir");
    await change(makeItem("B"));
    await advance(2000);
    tab("Automático");
    await change(makeItem("C"));
    await advance(2000);
    // the Automático panel makes its own settings calls; only analyze-message matters here
    expect(
      fetchMock.mock.calls.filter((c) => String(c[0]).includes("analyze-message"))
    ).toHaveLength(0);
  });
  it("(g) empty body, null item and a token that needs interaction send nothing; the sign-in text shows once", async () => {
    setup(makeItem("   "), true);
    render(<App />);
    await advance(2000);
    await change(null);
    await advance(2000);
    expect(fetchMock).not.toHaveBeenCalled();
    (getTokenSilent as any).mockResolvedValue({ kind: "interaction" });
    await change(makeItem("corpo"));
    await advance(800);
    await advance(5000);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(getTokenSilent).toHaveBeenCalledTimes(1);
    expect(screen.getAllByText("Inicie sessão (em baixo) para usar a IA.")).toHaveLength(1);
  });
  it("(h) a manual request pending: no auto request for that item", async () => {
    setup(makeItem("corpo"), true);
    const calls = hangingFetch();
    render(<App />);
    tab("Traduzir");
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Traduzir" }));
    });
    expect(calls).toHaveLength(1);
    tab("Resumo");
    await advance(3000);
    expect(calls).toHaveLength(1);
  });
  it("(i) with the preference off nothing is sent", async () => {
    setup(makeItem("A"), false);
    render(<App />);
    await advance(2000);
    await change(makeItem("B"));
    await advance(2000);
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe("errors and pause (F-003/AC-44)", () => {
  const pauseCases: [number, unknown, RegExp][] = [
    [429, { ok: false, error: "cap_reached" }, /fica em pausa até amanhã/],
    [401, { ok: false, error: "unauthorized" }, /A sessão expirou.*pausa até reabrir o painel/],
    [503, { ok: false, error: "disabled" }, /desligada.*pausa até reabrir o painel/],
    [503, { ok: false, error: "ai_not_configured" }, /configurada.*pausa até reabrir o painel/],
  ];
  it.each(pauseCases)(
    "status %s pauses auto mode for the session, once",
    async (status, body, text) => {
      setup(makeItem("A"), true);
      fetchMock.mockResolvedValue(json(status, body));
      render(<App />);
      await advance(800);
      expect(screen.getAllByRole("alert")).toHaveLength(1);
      expect(screen.getByRole("alert").textContent).toMatch(text);
      for (const b of ["B", "C", "D"]) {
        await change(makeItem(b));
        await advance(2000);
      }
      expect(fetchMock).toHaveBeenCalledTimes(1);
      expect((screen.getByRole("switch") as HTMLInputElement).checked).toBe(true);
      expect(rs.set).not.toHaveBeenCalled();
      fetchMock.mockResolvedValue(summaryOk());
      await act(async () => {
        fireEvent.click(screen.getByRole("button", { name: "Resumir" }));
      });
      expect(fetchMock).toHaveBeenCalledTimes(2);
    }
  );
  it("a retryable error has no automatic retry and the next item sends one request", async () => {
    setup(makeItem("A"), true);
    fetchMock.mockResolvedValue(json(502, { ok: false, error: "ai_output_invalid" }));
    render(<App />);
    await advance(800);
    expect(screen.getByRole("alert").textContent).toContain("A IA devolveu uma resposta inválida");
    await advance(60000);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    await change(makeItem("B"));
    await advance(800);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(sentBodies()[1].body).toBe("B");
  });
  it("never has more than one request in flight", async () => {
    setup(makeItem("A"), true);
    const calls = hangingFetch();
    render(<App />);
    await advance(800);
    expect(screen.getByRole("button", { name: "A resumir…" })).toBeTruthy();
    expect((screen.getByRole("button", { name: "A resumir…" }) as HTMLButtonElement).disabled).toBe(
      true
    );
    expect(calls).toHaveLength(1);
  });
});
