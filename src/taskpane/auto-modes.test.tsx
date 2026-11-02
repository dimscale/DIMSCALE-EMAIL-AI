import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

// F-003 amendment 2026-11-02 (i), AC-59 to AC-70 and the approved overrides (Q-U17 translate follows the Select,
// Q-U18 dialog once per switch, Q-U20 no-reply senders skipped, Q-U22 old autoSummary = true is consent).
// REAL config (flag true). Mocks only: Office, auth, fetch. Fake timers.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
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
const INJECT = "<img src=x onerror=alert(1)> &amp; fim";
const RESP: Record<string, unknown> = {
  summarize: { ok: true, action: "summarize", summary: SUMMARY },
  reply: { ok: true, action: "reply", reply: "minha resposta" },
  translate: { ok: true, action: "translate", translation: { subject: "S", body: "traduzido" } },
};
const ARIA = {
  summarize: "Resumo automático",
  reply: "Resposta automática",
  translate: "Tradução automática",
};
const KEY = { summarize: "autoSummary", reply: "autoReply", translate: "autoTranslate" };
type Act = "summarize" | "reply" | "translate";

const handlers = new Set<() => void>();
let office: any;
let rs: { get: any; set: any; saveAsync: any };
let item: any;
let fetchMock: ReturnType<typeof vi.fn>;
let inFlight = 0;
let maxInFlight = 0;

const makeItem = (body: string, o: object = {}) => ({
  subject: "Assunto",
  from: { displayName: "Ana Silva", emailAddress: "ana@exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  displayReplyFormAsync: vi.fn(),
  displayReplyForm: vi.fn(),
  displayReplyAllForm: vi.fn(),
  displayReplyAllFormAsync: vi.fn(),
  ...o,
});
const setup = (it: unknown, store: Record<string, unknown> = {}) => {
  handlers.clear();
  item = it;
  rs = {
    get: vi.fn((k: string) => store[k]),
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
        displayNewMessageForm: vi.fn(),
        addHandlerAsync: vi.fn((e: string, h: () => void) => {
          if (e === "itemChanged") handlers.add(h);
        }),
        removeHandlerAsync: vi.fn(),
      },
    },
  };
  (globalThis as any).Office = office;
};
const on = (...a: Act[]) => Object.fromEntries(a.map((x) => [KEY[x], true]));
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
const change = async (it: unknown) => {
  office.context.mailbox.item = it;
  item = it;
  await act(async () => handlers.forEach((h) => h()));
};
const tab = (name: string) => fireEvent.click(screen.getByRole("tab", { name }));
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name }));
  });
};
const analyzeCalls = () =>
  fetchMock.mock.calls.filter((c) => String(c[0]).includes("analyze-message"));
const sent = () => analyzeCalls().map((c) => JSON.parse(c[1].body));
const actions = () => sent().map((b) => b.action);
const answerByAction = () =>
  fetchMock.mockImplementation(async (_u: string, init: { body: string }) =>
    json(200, RESP[JSON.parse(init.body).action])
  );
/* A fetch that stays pending until resolved, and rejects on abort like the real one. */
const hangingFetch = () => {
  const calls: { signal: AbortSignal; action: string; resolve: (v: unknown) => void }[] = [];
  fetchMock.mockImplementation(
    (_u: string, init: { signal: AbortSignal; body: string }) =>
      new Promise((resolve, reject) => {
        calls.push({ signal: init.signal, action: JSON.parse(init.body).action, resolve });
        init.signal.addEventListener("abort", () => reject(new Error("aborted")));
      })
  );
  return calls;
};
const sw = (a: Act) => screen.getByRole("switch", { name: ARIA[a] }) as HTMLInputElement;
const turnOn = async (a: Act) => {
  await act(async () => {
    fireEvent.click(sw(a));
  });
};
const ligar = async () => {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name: "Ligar" }));
  });
};

beforeEach(() => {
  vi.useFakeTimers();
  inFlight = 0;
  maxInFlight = 0;
  fetchMock = vi.fn();
  answerByAction();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.clearAllMocks();
});

describe("three switches (F-003/AC-59, AC-72)", () => {
  it("each tool tab shows exactly its own switch first in the options, with an action-specific name; Lista shows none of them", () => {
    setup(makeItem("corpo"));
    render(<App />);
    for (const [name, a, next] of [
      ["Resumo", "summarize", "Resumir"],
      ["Resposta", "reply", "Idioma"],
      ["Traduzir", "translate", "Traduzir para"],
    ] as const) {
      tab(name);
      const switches = screen.getAllByRole("switch"); // the other cards are hidden, out of the tree
      expect(switches).toHaveLength(1);
      expect(switches[0]).toBe(sw(a));
      expect((switches[0] as HTMLInputElement).labels?.[0]?.textContent).toBe("Automático");
      const after =
        next === "Resumir"
          ? screen.getByRole("button", { name: next })
          : screen.getByLabelText(next);
      expect(
        switches[0].compareDocumentPosition(after) & Node.DOCUMENT_POSITION_FOLLOWING
      ).toBeTruthy();
    }
    tab("Lista");
    for (const a of ["summarize", "reply", "translate"] as const)
      expect(screen.queryByRole("switch", { name: ARIA[a] })).toBeNull();
  });
  it("the three switches are independent and storage keys are their own", async () => {
    setup(makeItem("corpo"), { ...on("summarize"), autoReplyConsent: true });
    render(<App />);
    expect(sw("summarize").checked).toBe(true);
    tab("Resposta");
    expect(sw("reply").checked).toBe(false);
    await turnOn("reply"); // consent remembered: no dialog
    expect(screen.queryByRole("dialog")).toBeNull();
    expect(sw("reply").checked).toBe(true);
    expect(rs.set.mock.calls).toEqual([
      ["autoReply", true],
      ["autoReplyConsent", true],
    ]);
    tab("Traduzir");
    expect(sw("translate").checked).toBe(false);
    tab("Resumo");
    expect(sw("summarize").checked).toBe(true);
  });
  it("shows the 'ligado' line of its own action", () => {
    setup(makeItem("corpo"), on("summarize", "reply", "translate"));
    render(<App />);
    expect(screen.getByText(/^Resumo automático ligado/)).toBeTruthy();
    expect(screen.getByText(/^Resposta automática ligada/)).toBeTruthy();
    expect(screen.getByText(/^Tradução automática ligada/)).toBeTruthy();
  });
});

describe("default off and consent dialog (F-003/AC-60, AC-61, Q-U18, Q-U22)", () => {
  it("a fresh user: render, ItemChanged, tab switches and sign-in make zero requests", async () => {
    setup(makeItem("A"));
    render(<App />);
    await advance(1000);
    for (const t of ["Resposta", "Traduzir", "Lista", "Resumo"]) tab(t);
    await change(makeItem("B"));
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(0);
  });
  it("an old autoSummary = true turns only the summary on and counts as consent (no dialog later)", async () => {
    setup(makeItem("A"), { autoSummary: true });
    render(<App />);
    expect(sw("summarize").checked).toBe(true);
    await advance(800);
    expect(actions()).toEqual(["summarize"]);
    await turnOn("summarize"); // off
    await turnOn("summarize"); // on again: no dialog
    expect(screen.queryByRole("dialog")).toBeNull();
    tab("Resposta");
    expect(sw("reply").checked).toBe(false);
  });
  it.each([
    ["Resumo", "summarize", "Ligar o resumo automático?", "sem clicar em Resumir"],
    [
      "Resposta",
      "reply",
      "Ligar a resposta automática?",
      "A resposta só aparece no painel: nada é enviado nem aberto automaticamente.",
    ],
    [
      "Traduzir",
      "translate",
      "Ligar a tradução automática?",
      "(sem o nome nem o domínio do remetente)",
    ],
  ] as const)(
    "%s: the first time opens its own dialog; Cancelar and Escape leave it off; Ligar saves and remembers",
    async (tabName, a, title, bodyPart) => {
      setup(makeItem("corpo"));
      render(<App />);
      tab(tabName);
      await turnOn(a);
      expect(screen.getByText(title)).toBeTruthy();
      expect(screen.getByRole("dialog").textContent).toContain(bodyPart);
      fireEvent.click(screen.getByRole("button", { name: "Cancelar" }));
      await advance(2000);
      expect(sw(a).checked).toBe(false);
      await turnOn(a);
      fireEvent.keyDown(screen.getByRole("dialog"), { key: "Escape" });
      await advance(2000);
      expect(sw(a).checked).toBe(false);
      expect(rs.set).not.toHaveBeenCalled();
      expect(analyzeCalls()).toHaveLength(0); // nothing leaves before "Ligar"
      await turnOn(a);
      await ligar();
      expect(sw(a).checked).toBe(true);
      expect(rs.set.mock.calls).toEqual([
        [KEY[a], true],
        [`${KEY[a]}Consent`, true],
      ]);
      expect(rs.saveAsync).toHaveBeenCalledTimes(1);
      await turnOn(a); // off: no dialog
      expect(screen.queryByRole("dialog")).toBeNull();
      await turnOn(a); // on again: consent remembered, no dialog
      expect(screen.queryByRole("dialog")).toBeNull();
      expect(sw(a).checked).toBe(true);
    }
  );
  it("a failed save keeps the choice for the session and shows the notice under its own switch only", async () => {
    setup(makeItem("corpo"));
    rs.saveAsync.mockImplementation((cb: Cb) => cb({ status: "failed" }));
    render(<App />);
    tab("Resposta");
    await turnOn("reply");
    await ligar();
    expect(sw("reply").checked).toBe(true);
    expect(screen.getAllByText(/Não foi possível guardar a preferência/)).toHaveLength(1);
    tab("Resumo");
    expect(screen.getAllByText(/Não foi possível guardar a preferência/)).toHaveLength(1); // not under Resumo's switch
  });
});

describe("triggers, visible tab only (F-003/AC-63)", () => {
  it("(a) on load only the summary can run; with only reply and translate on, nothing", async () => {
    setup(makeItem("A"), on("summarize", "reply", "translate"));
    render(<App />);
    await advance(799);
    expect(analyzeCalls()).toHaveLength(0);
    await advance(1);
    await advance(5000);
    expect(actions()).toEqual(["summarize"]);
    cleanup();
    fetchMock.mockClear();
    setup(makeItem("A"), on("reply", "translate"));
    render(<App />);
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(0);
  });
  it("(b) ItemChanged on Resposta sends one reply for B, no summarize and no translate", async () => {
    setup(makeItem("A", { itemId: "a" }), on("summarize", "reply", "translate"));
    render(<App />);
    tab("Resposta");
    await advance(800);
    expect(actions()).toEqual(["reply"]);
    await change(makeItem("B", { itemId: "b" }));
    await advance(5000);
    expect(sent().map((b) => [b.action, b.body])).toEqual([
      ["reply", "A"],
      ["reply", "B"],
    ]);
  });
  it("(c) switching tabs sends the entered tab's action once; back to a tab with none yet sends it", async () => {
    setup(makeItem("A"), on("summarize", "translate"));
    render(<App />);
    tab("Traduzir"); // before the summary debounce ends: its timer is cancelled
    await advance(800);
    expect(actions()).toEqual(["translate"]);
    tab("Resumo");
    await advance(800);
    expect(actions()).toEqual(["translate", "summarize"]);
    tab("Traduzir");
    tab("Resumo");
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(2);
  });
  it("(d) turning the switch on with Ligar on its own tab sends one request after 800 ms", async () => {
    setup(makeItem("A"));
    render(<App />);
    tab("Resposta");
    await turnOn("reply");
    await ligar();
    await advance(799);
    expect(analyzeCalls()).toHaveLength(0);
    await advance(1);
    expect(actions()).toEqual(["reply"]);
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(1);
  });
  it("(e) A, B, C within 800 ms send one request, for C", async () => {
    setup(makeItem("A"), on("reply"));
    render(<App />);
    tab("Resposta");
    await advance(300);
    await change(makeItem("B"));
    await advance(300);
    await change(makeItem("C"));
    await advance(799);
    expect(analyzeCalls()).toHaveLength(0);
    await advance(1);
    expect(sent().map((b) => b.body)).toEqual(["C"]);
  });
  it("(f) Lista selected and (g) a visible tab whose switch is off: nothing", async () => {
    setup(makeItem("A"), on("summarize", "reply", "translate"));
    render(<App />);
    tab("Lista");
    await advance(5000);
    await change(makeItem("B"));
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(0);
    cleanup();
    setup(makeItem("A"), on("summarize", "translate"));
    render(<App />);
    tab("Resposta");
    await advance(5000);
    await change(makeItem("B"));
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(0);
  });
});

describe("one request per email per tab, waiting, one in flight (F-003/AC-64, AC-65, AC-69)", () => {
  it("a revisit sends nothing; the same email on another tab is a separate pair", async () => {
    setup(makeItem("A", { itemId: "a" }), on("reply", "translate"));
    render(<App />);
    tab("Resposta");
    await advance(800);
    await change(makeItem("B", { itemId: "b" }));
    await advance(800);
    await change(makeItem("A", { itemId: "a" }));
    await advance(5000);
    expect(sent().map((b) => b.body)).toEqual(["A", "B"]);
    tab("Traduzir");
    await advance(800);
    expect(sent().map((b) => [b.action, b.body])).toEqual([
      ["reply", "A"],
      ["reply", "B"],
      ["translate", "A"],
    ]);
  });
  it("a manual click counts as the attempt for that tab and email", async () => {
    setup(makeItem("A"), on("reply"));
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta"); // before the debounce
    await advance(5000);
    expect(actions()).toEqual(["reply"]);
  });
  it("a pending manual request makes auto wait, not skip: it goes out 800 ms after the pending one ends", async () => {
    setup(makeItem("A"), on("reply"));
    const calls = hangingFetch();
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    tab("Resposta");
    await advance(3000);
    expect(calls).toHaveLength(1);
    await act(async () => calls[0].resolve(json(200, RESP.translate)));
    await advance(799);
    expect(calls).toHaveLength(1);
    await advance(1);
    expect(calls.map((c) => c.action)).toEqual(["translate", "reply"]);
  });
  it("a pending request ending in an error does not retry", async () => {
    setup(makeItem("A"), on("summarize"));
    fetchMock.mockResolvedValue(json(502, { ok: false, error: "ai_output_invalid" }));
    render(<App />);
    await advance(60000);
    expect(analyzeCalls()).toHaveLength(1);
  });
  it("leaving a tab does not abort its in-flight request; the result is there on return", async () => {
    setup(makeItem("A"), on("translate"));
    const calls = hangingFetch();
    render(<App />);
    tab("Traduzir");
    await advance(800);
    tab("Resumo");
    await advance(100);
    expect(calls[0].signal.aborted).toBe(false);
    await act(async () => calls[0].resolve(json(200, RESP.translate)));
    tab("Traduzir");
    expect(screen.getByText("traduzido")).toBeTruthy();
  });
  it("turning a switch off aborts only that action's in-flight auto request, not a manual one", async () => {
    setup(makeItem("A"), on("reply", "translate"));
    const calls = hangingFetch();
    render(<App />);
    tab("Resposta");
    await advance(800);
    expect(calls).toHaveLength(1);
    await turnOn("reply"); // off
    expect(calls[0].signal.aborted).toBe(true);
    await act(async () => calls[0].resolve(json(200, RESP.reply)));
    expect(screen.queryByText("minha resposta")).toBeNull();
    expect(screen.queryByRole("alert")).toBeNull();
    // a manual translate request survives turning the translate switch off
    tab("Traduzir");
    await click("Traduzir");
    await turnOn("translate"); // off
    expect(calls[calls.length - 1].action).toBe("translate");
    expect(calls[calls.length - 1].signal.aborted).toBe(false);
  });
  it("all switches on, every trigger exercised: never more than one request in flight", async () => {
    setup(makeItem("A", { itemId: "a" }), on("summarize", "reply", "translate"));
    let aborts = 0;
    fetchMock.mockImplementation((_u: string, init: { signal: AbortSignal; body: string }) => {
      inFlight++;
      maxInFlight = Math.max(maxInFlight, inFlight);
      return new Promise((resolve, reject) => {
        let done = false;
        const settle = () => {
          if (!done) {
            done = true;
            inFlight--;
          }
        };
        init.signal.addEventListener("abort", () => {
          aborts++;
          settle();
          reject(new Error("aborted"));
        });
        setTimeout(() => {
          settle();
          resolve(json(200, RESP[JSON.parse(init.body).action]));
        }, 300);
      });
    });
    render(<App />);
    for (const id of ["a", "b", "c"]) {
      for (const t of ["Resumo", "Resposta", "Traduzir", "Resumo", "Traduzir"]) {
        tab(t);
        await advance(900);
      }
      tab("Resposta");
      await advance(850); // a request is in flight (300 ms server) or just done when the email changes
      await change(makeItem(id.toUpperCase(), { itemId: id + "2" }));
    }
    await advance(5000);
    expect(analyzeCalls().length).toBeGreaterThan(5);
    expect(maxInFlight).toBe(1);
    expect(aborts).toBeGreaterThan(0); // the abort path was really exercised
  });
  it("manual buttons: disabled while any request is pending, and a click after an auto result sends a new request", async () => {
    setup(makeItem("A"), on("summarize"));
    const calls = hangingFetch();
    render(<App />);
    await advance(800);
    for (const n of ["Sugerir resposta", "Traduzir"])
      expect(
        (screen.getByRole("button", { name: n, hidden: true }) as HTMLButtonElement).disabled
      ).toBe(true);
    await act(async () => calls[0].resolve(json(200, RESP.summarize)));
    answerByAction();
    await click("Resumir");
    expect(actions()).toEqual(["summarize", "summarize"]);
  });
});

describe("payloads (F-003/AC-66, Q-U17)", () => {
  it("auto reply carries the Idioma and Tom current when the timer fires, and no instruction or draft", async () => {
    setup(makeItem("corpo"), on("reply"));
    render(<App />);
    tab("Resposta");
    fireEvent.change(screen.getByLabelText("Idioma"), { target: { value: "en" } });
    fireEvent.change(screen.getByLabelText("Tom"), { target: { value: "formal" } });
    await advance(800);
    expect(sent()).toEqual([
      {
        action: "reply",
        subject: "Assunto",
        fromName: "Ana Silva",
        fromDomain: "exemplo.pt",
        body: "corpo",
        language: "en",
        tone: "formal",
      },
    ]);
  });
  it("auto translate follows the 'Traduzir para' value and sends no sender data", async () => {
    setup(makeItem("corpo"), on("translate"));
    render(<App />);
    tab("Traduzir");
    fireEvent.change(screen.getByLabelText("Traduzir para"), { target: { value: "pt" } });
    await advance(800);
    expect(sent()).toEqual([
      {
        action: "translate",
        subject: "Assunto",
        fromName: "",
        fromDomain: "",
        body: "corpo",
        targetLanguage: "pt",
      },
    ]);
    const raw = analyzeCalls()[0][1].body as string;
    expect(raw).not.toContain("Ana");
    expect(raw).not.toContain("exemplo");
  });
  it("auto translate with the default Select sends targetLanguage auto", async () => {
    setup(makeItem("corpo"), on("translate"));
    render(<App />);
    tab("Traduzir");
    await advance(800);
    expect(sent()[0].targetLanguage).toBe("auto");
  });
});

describe("when auto does not run (F-003/AC-67)", () => {
  it("empty body (note shows), body error and no item send nothing on Resposta", async () => {
    setup(makeItem("   "), on("reply"));
    render(<App />);
    tab("Resposta");
    await advance(2000);
    expect(
      screen.getAllByText("Esta mensagem não tem texto para analisar.").length
    ).toBeGreaterThan(0);
    await change(null);
    await advance(2000);
    await change(
      makeItem("corpo", {
        body: { getAsync: (_t: string, cb: Cb) => cb({ status: "failed", value: undefined }) },
      })
    );
    await advance(2000);
    expect(analyzeCalls()).toHaveLength(0);
  });
  it("a token that needs interaction: no request, the text shows once, no loop, and no request after signing in until the next trigger", async () => {
    setup(makeItem("corpo"), on("reply"));
    (getTokenSilent as any).mockResolvedValue({ kind: "interaction" });
    render(<App />);
    tab("Resposta");
    await advance(800);
    await advance(5000);
    expect(getTokenSilent).toHaveBeenCalledTimes(1);
    expect(screen.getAllByText("Inicie sessão (em baixo) para usar a IA.")).toHaveLength(1);
    (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(0);
  });
});

describe("no-reply senders (Q-U20)", () => {
  it.each(["noreply@exemplo.pt", "No-Reply+x@exemplo.pt", "mailer-daemon@exemplo.pt"])(
    "%s: auto reply and auto translate send nothing; auto summary runs; manual buttons work",
    async (from) => {
      const mk = (b: string) =>
        makeItem(b, { from: { displayName: "Sistema", emailAddress: from } });
      setup(mk("A"), on("summarize", "reply", "translate"));
      render(<App />);
      await advance(800);
      expect(actions()).toEqual(["summarize"]);
      tab("Resposta");
      await advance(3000);
      tab("Traduzir");
      await advance(3000);
      await change(mk("B"));
      await advance(3000);
      expect(actions()).toEqual(["summarize"]);
      await click("Traduzir");
      expect(actions()).toEqual(["summarize", "translate"]);
    }
  );
  it("a normal sender with noreply inside the local part is not skipped", async () => {
    setup(
      makeItem("A", { from: { displayName: "X", emailAddress: "noreplyfan@exemplo.pt" } }),
      on("reply")
    );
    render(<App />);
    tab("Resposta");
    await advance(800);
    expect(actions()).toEqual(["reply"]);
  });
});

describe("shared pause (F-003/AC-68, Q-U23)", () => {
  const cases: [string, number, unknown, RegExp, RegExp][] = [
    [
      "cap",
      429,
      { ok: false, error: "cap_reached" },
      /fica em pausa até amanhã/,
      /^Modo automático em pausa até amanhã \(limite diário atingido\)\.$/,
    ],
    [
      "401",
      401,
      { ok: false, error: "unauthorized" },
      /A sessão expirou.*pausa até reabrir o painel/,
      /^Modo automático em pausa até reabrir o painel\.$/,
    ],
    [
      "disabled",
      503,
      { ok: false, error: "disabled" },
      /desligada.*pausa até reabrir o painel/,
      /^Modo automático em pausa até reabrir o painel\.$/,
    ],
    [
      "not configured",
      503,
      { ok: false, error: "ai_not_configured" },
      /configurada.*pausa até reabrir o painel/,
      /^Modo automático em pausa até reabrir o painel\.$/,
    ],
    [
      "rejected",
      503,
      { ok: false, error: "ai_unavailable", retryable: false, why: "http_401" },
      /recusou o pedido.*pausa até reabrir o painel/,
      /^Modo automático em pausa até reabrir o painel\.$/,
    ],
  ];
  it.each(cases)(
    "%s on an auto reply pauses every tab: alert on the failing card, status line on the other card, zero more auto requests, manual still works",
    async (_n, status, body, alertRe, lineRe) => {
      setup(makeItem("A"), on("reply", "translate"));
      fetchMock.mockResolvedValue(json(status, body));
      render(<App />);
      tab("Resposta");
      await advance(800);
      expect(screen.getAllByRole("alert")).toHaveLength(1);
      expect(screen.getByRole("alert").textContent).toMatch(alertRe);
      expect(screen.getAllByText(lineRe)).toHaveLength(1); // the translate card only
      expect(screen.getByText(lineRe).closest("[role=status]")).toBeTruthy();
      for (const b of ["B", "C", "D"]) {
        await change(makeItem(b));
        for (const t of ["Traduzir", "Resumo", "Resposta"]) {
          tab(t);
          await advance(2000);
        }
      }
      expect(analyzeCalls()).toHaveLength(1);
      // the alert is gone after the email change: the failing card now shows the pause line too
      expect(screen.queryByRole("alert")).toBeNull();
      expect(screen.getAllByText(lineRe)).toHaveLength(2);
      expect(sw("reply").checked).toBe(true);
      expect(rs.set).not.toHaveBeenCalled();
      answerByAction();
      await click("Sugerir resposta");
      tab("Traduzir");
      await click("Traduzir");
      expect(actions()).toEqual(["reply", "reply", "translate"]);
    }
  );
  it("a manual request ending in a pause code does not pause auto mode", async () => {
    setup(makeItem("A"), on("reply"));
    fetchMock.mockResolvedValueOnce(json(429, { ok: false, error: "cap_reached" }));
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    expect(screen.getByRole("alert").textContent).toContain("Tente amanhã");
    tab("Resposta");
    await advance(800);
    expect(actions()).toEqual(["translate", "reply"]);
  });
  it("a retryable error and invalid do not pause; the next email sends; no retry in 60 s", async () => {
    setup(makeItem("A"), on("translate"));
    fetchMock.mockResolvedValueOnce(json(502, { ok: false, error: "ai_output_invalid" }));
    render(<App />);
    tab("Traduzir");
    await advance(800);
    expect(screen.getByRole("alert").textContent).toContain("inválida");
    expect(screen.queryByText(/em pausa/)).toBeNull();
    await advance(60000);
    expect(analyzeCalls()).toHaveLength(1);
    await change(makeItem("B"));
    await advance(800);
    expect(sent().map((b) => b.body)).toEqual(["A", "B"]);
  });
  it("the pause is per pane session: a new App mount starts unpaused", async () => {
    setup(makeItem("A"), on("reply"));
    fetchMock.mockResolvedValueOnce(json(429, { ok: false, error: "cap_reached" }));
    render(<App />);
    tab("Resposta");
    await advance(800);
    cleanup();
    setup(makeItem("A"), on("reply"));
    render(<App />);
    tab("Resposta");
    await advance(800);
    expect(actions()).toEqual(["reply", "reply"]);
  });
});

describe("nothing is sent or opened automatically (F-003/AC-70)", () => {
  it("all switches on and every trigger exercised: no reply form, no clipboard, no compose call; the reply renders as text", async () => {
    const writeText = vi.fn();
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    const made: any[] = [];
    const mk = (b: string, id: string) => {
      const m = makeItem(b, { itemId: id });
      made.push(m);
      return m;
    };
    setup(mk("A", "a"), on("summarize", "reply", "translate"));
    fetchMock.mockImplementation(async (_u: string, init: { body: string }) => {
      const a = JSON.parse(init.body).action;
      return json(200, a === "reply" ? { ok: true, action: "reply", reply: INJECT } : RESP[a]);
    });
    const { container } = render(<App />);
    for (const t of ["Resumo", "Resposta", "Traduzir", "Lista", "Resposta"]) {
      tab(t);
      await advance(900);
    }
    expect(actions()).toEqual(["summarize", "reply", "translate"]);
    expect(screen.getAllByText(INJECT).length).toBeGreaterThan(0); // text node, not markup
    expect(container.querySelector("img")).toBeNull();
    await change(mk("B", "b"));
    await advance(900);
    expect(analyzeCalls()).toHaveLength(4); // reply for B on the visible tab
    for (const it of made) {
      expect(it.displayReplyForm).not.toHaveBeenCalled();
      expect(it.displayReplyFormAsync).not.toHaveBeenCalled();
      expect(it.displayReplyAllForm).not.toHaveBeenCalled();
      expect(it.displayReplyAllFormAsync).not.toHaveBeenCalled();
    }
    expect(office.context.mailbox.displayNewMessageForm).not.toHaveBeenCalled();
    expect(writeText).not.toHaveBeenCalled();
  });
});
