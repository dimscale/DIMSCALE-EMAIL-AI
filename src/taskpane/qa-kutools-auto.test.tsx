import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

// QA (independent): auto mode adversarial cases (F-003/AC-42 to AC-44), availability flag forced ON. Fake timers. Mocks only.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
vi.mock("./config", async (orig) => ({
  ...(await orig<Record<string, unknown>>()),
  AUTO_SUMMARY_AVAILABLE: true,
}));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
const SUMMARY = {
  points: ["um", "dois", "tres"],
  type: "request",
  priority: "high",
  tone: "neutral",
  deadline: null,
  requestedAction: null,
};
const DIALOG_BODY =
  "Com o resumo automático, sempre que abrir um e-mail com este painel no separador Resumo, o assunto, o nome e domínio do remetente e o texto do e-mail são enviados à OpenAI para gerar o resumo, sem clicar em Resumir. Cada resumo conta para o seu limite diário de 50 pedidos de IA. Pode desligar a qualquer momento.";
const ON_NOTE =
  "Resumo automático ligado: cada e-mail que abrir com o painel neste separador é enviado à OpenAI.";

const handlers = new Set<() => void>();
let office: any;
let rs: { get: any; set: any; saveAsync: any };
let fetchMock: ReturnType<typeof vi.fn>;
// in flight = called and not yet settled (resolved, rejected or aborted)
let inFlight = 0;
let maxInFlight = 0;

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
const toggle = async () => {
  await act(async () => {
    fireEvent.click(screen.getByRole("switch"));
  });
};
const analyzeCalls = () =>
  fetchMock.mock.calls.filter((c) => String(c[0]).includes("analyze-message"));
const sentBodies = () => analyzeCalls().map((c) => JSON.parse(c[1].body).body);
const ok = () => json(200, { ok: true, action: "summarize", summary: SUMMARY });

/* Every call is tracked; it stops being "in flight" when it resolves or when its signal aborts (a real fetch rejects). */
const trackedFetch = (answerAfterMs: number) =>
  vi.fn((_u: string, init: { signal: AbortSignal }) => {
    inFlight++;
    maxInFlight = Math.max(maxInFlight, inFlight);
    return new Promise((resolve, reject) => {
      let settled = false;
      const done = () => {
        if (!settled) {
          settled = true;
          inFlight--;
        }
      };
      init.signal.addEventListener("abort", () => {
        done();
        reject(new Error("aborted"));
      });
      setTimeout(() => {
        done();
        resolve(ok());
      }, answerAfterMs);
    });
  });

beforeEach(() => {
  vi.useFakeTimers();
  inFlight = 0;
  maxInFlight = 0;
  fetchMock = vi.fn(async () => ok());
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.clearAllMocks();
});

describe("switch off and on again (F-003/AC-42, AC-43)", () => {
  it("switch off then on again for the same item sends no second request for that item", async () => {
    setup(makeItem("A"), true);
    render(<App />);
    await advance(800);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    await toggle(); // off, no dialog
    await toggle(); // on: dialog
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    await advance(5000);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(sentBodies()).toEqual(["A"]);
  });

  it("an in-flight auto request aborted by switching off is not sent again when switching back on", async () => {
    setup(makeItem("A"), true);
    const calls: { signal: AbortSignal }[] = [];
    fetchMock.mockImplementation((_u: string, init: { signal: AbortSignal }) => {
      calls.push({ signal: init.signal });
      return new Promise((_r, rej) =>
        init.signal.addEventListener("abort", () => rej(new Error("x")))
      );
    });
    render(<App />);
    await advance(800);
    expect(calls).toHaveLength(1);
    await toggle();
    expect(calls[0].signal.aborted).toBe(true);
    await toggle();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    await advance(5000);
    expect(calls).toHaveLength(1);
  });

  it("the dialog text and the 'ligado' line match the spec text exactly", async () => {
    setup(makeItem("A"), false);
    render(<App />);
    await toggle();
    expect(screen.getByText(DIALOG_BODY)).toBeTruthy();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    expect(screen.getByText(ON_NOTE)).toBeTruthy();
  });
});

describe("saved preference edge cases (F-003/AC-42)", () => {
  it("a failed save keeps auto mode running for the next item (in-memory for the session)", async () => {
    setup(makeItem("A"), false, "failed");
    render(<App />);
    await toggle();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    expect(screen.getByText(/Não foi possível guardar a preferência/)).toBeTruthy();
    await advance(800);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    await change(makeItem("B"));
    await advance(800);
    expect(sentBodies()).toEqual(["A", "B"]);
  });

  it("no roamingSettings object at all: off, and switching on still runs for this session (save notice shown)", async () => {
    setup(makeItem("A"), false);
    office.context.roamingSettings = undefined;
    render(<App />);
    expect((screen.getByRole("switch") as HTMLInputElement).checked).toBe(false);
    await toggle();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    expect(screen.getByText(/Não foi possível guardar a preferência/)).toBeTruthy();
    await advance(800);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("roamingSettings.get throws: the switch starts off and nothing is sent", async () => {
    setup(makeItem("A"), false);
    rs.get.mockImplementation(() => {
      throw new Error("denied");
    });
    render(<App />);
    expect((screen.getByRole("switch") as HTMLInputElement).checked).toBe(false);
    await change(makeItem("B"));
    await advance(5000);
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe("pause holds for the pane session (F-003/AC-44)", () => {
  it("after a 429 pause, switching off and on again does not resume auto mode for the session", async () => {
    setup(makeItem("A"), true);
    fetchMock.mockResolvedValue(json(429, { ok: false, error: "cap_reached" }));
    render(<App />);
    await advance(800);
    expect(screen.getByRole("alert").textContent).toContain("fica em pausa até amanhã");
    await toggle();
    await toggle();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
    });
    await change(makeItem("B"));
    await advance(5000);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("a pause does not block a manual click on the next item", async () => {
    setup(makeItem("A"), true);
    fetchMock.mockResolvedValueOnce(json(429, { ok: false, error: "cap_reached" }));
    render(<App />);
    await advance(800);
    await change(makeItem("B"));
    await advance(5000);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    fetchMock.mockResolvedValueOnce(ok());
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Resumir" }));
    });
    expect(sentBodies()).toEqual(["A", "B"]);
  });

  it("a retryable error: leaving Resumo and coming back does not retry the same item", async () => {
    setup(makeItem("A"), true);
    fetchMock.mockResolvedValue(json(503, { ok: false, error: "ai_unavailable", retryable: true }));
    render(<App />);
    await advance(800);
    expect(screen.getByRole("alert")).toBeTruthy();
    tab("Resposta");
    tab("Resumo");
    await advance(5000);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe("tab and entry interplay (F-003/AC-43, F-012/AC-50)", () => {
  it("items changed while on Automático: one auto request after 800 ms when Resumo is selected again, for the current item only", async () => {
    setup(makeItem("A"), true);
    render(<App />);
    tab("Automático");
    await advance(2000); // A is not sent while on Automático
    expect(analyzeCalls()).toHaveLength(0);
    await change(makeItem("B"));
    await change(makeItem("C"));
    await advance(2000);
    expect(analyzeCalls()).toHaveLength(0);
    tab("Resumo");
    await advance(799);
    expect(analyzeCalls()).toHaveLength(0);
    await advance(1);
    expect(sentBodies()).toEqual(["C"]);
  });

  it("entry page with auto on: the loaded item gets exactly one request (the entry), not a second one after 800 ms; the next item gets one auto request", async () => {
    setup(makeItem("A"), true);
    render(<App summaryEntry />);
    await advance(0);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    await advance(5000);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    await change(makeItem("B"));
    await advance(800);
    expect(sentBodies()).toEqual(["A", "B"]);
  });
});

describe("at most one request in flight (F-003/AC-44), measured", () => {
  it("auto A -> B -> C with a 300 ms server and abort-honouring fetch: never more than one in flight", async () => {
    setup(makeItem("A"), true);
    fetchMock = trackedFetch(300);
    vi.stubGlobal("fetch", fetchMock);
    render(<App />);
    await advance(900); // A sent
    await change(makeItem("B"));
    await advance(300);
    await change(makeItem("C"));
    await advance(2000);
    expect(maxInFlight).toBe(1);
    // A was sent at 800 ms and aborted by the change to B; B's timer was cancelled by C; only C is sent after that
    expect(sentBodies()).toEqual(["A", "C"]);
  });

  it("manual click on A still in flight, then auto for B: the A request is aborted and at most one is live", async () => {
    setup(makeItem("A"), true);
    fetchMock = trackedFetch(5000);
    vi.stubGlobal("fetch", fetchMock);
    render(<App />);
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Resumir" }));
    });
    expect(inFlight).toBe(1);
    await change(makeItem("B"));
    await advance(800);
    expect(maxInFlight).toBe(1);
    expect(inFlight).toBe(1); // only B is live; A was aborted and settled
  });
});

describe("limitation: identity without itemId (F-003 AC-43c, documented)", () => {
  it("LIMITATION: with no itemId, returning to A in the same session sends A again (item identity falls back to the pane visit)", async () => {
    setup(makeItem("A"), true); // no itemId
    render(<App />);
    await advance(800);
    await change(makeItem("B"));
    await advance(800);
    await change(makeItem("A"));
    await advance(2000);
    // the spec says one request per item per pane session; without an id this is not met
    expect(sentBodies()).toEqual(["A", "B", "A"]);
  });
});

describe("auto: a token that arrives after the user left the email (F-003/AC-43, privacy)", () => {
  it("the 800 ms timer fired for A, the token is pending, the user moves to B: nothing for A; B gets its own request after 800 ms", async () => {
    setup(makeItem("A"), true);
    let resolveToken!: (v: unknown) => void;
    (getTokenSilent as any).mockImplementationOnce(() => new Promise((r) => (resolveToken = r)));
    render(<App />);
    await advance(800);
    expect(getTokenSilent).toHaveBeenCalledTimes(1);
    await change(makeItem("B"));
    await act(async () => resolveToken({ kind: "token", token: "TOK" }));
    await advance(799);
    expect(sentBodies()).toEqual([]);
    await advance(1);
    expect(sentBodies()).toEqual(["B"]);
  });
});
