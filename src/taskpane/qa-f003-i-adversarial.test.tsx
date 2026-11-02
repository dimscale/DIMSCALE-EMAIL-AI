import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

// QA adversarial checks for F-003 amendment 2026-11-02 (i) (AC-59 to AC-70, approved overrides) and F-012 AC-55.
// REAL config (flag true). Mocks only: Office (roamingSettings, mailbox), auth, fetch. Fake timers.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
type Act = "summarize" | "reply" | "translate";
const SUMMARY = {
  points: ["um", "dois"],
  type: "request",
  priority: "high",
  tone: "neutral",
  deadline: null,
  requestedAction: null,
};
const RESP: Record<string, unknown> = {
  summarize: { ok: true, action: "summarize", summary: SUMMARY },
  reply: { ok: true, action: "reply", reply: "minha resposta" },
  translate: { ok: true, action: "translate", translation: { subject: "S", body: "traduzido" } },
};
const ARIA: Record<Act, string> = {
  summarize: "Resumo automático",
  reply: "Resposta automática",
  translate: "Tradução automática",
};
const KEY: Record<Act, string> = {
  summarize: "autoSummary",
  reply: "autoReply",
  translate: "autoTranslate",
};
const CONSENT_TITLE = "Ligar a resposta automática?";
const SIGN_IN = "Inicie sessão (em baixo) para usar a IA.";

const handlers = new Set<() => void>();
let office: any;
let rs: { get: any; set: any; saveAsync: any };
let fetchMock: ReturnType<typeof vi.fn>;

const makeItem = (body: string, o: object = {}) => ({
  subject: "Assunto",
  from: { displayName: "Ana Silva", emailAddress: "ana@exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  displayReplyFormAsync: vi.fn(),
  displayReplyForm: vi.fn(),
  ...o,
});
const setup = (it: unknown, store: Record<string, unknown> = {}) => {
  handlers.clear();
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
        item: it,
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
  await act(async () => handlers.forEach((h) => h()));
};
const tab = (name: string) => fireEvent.click(screen.getByRole("tab", { name }));
const analyzeCalls = () =>
  fetchMock.mock.calls.filter((c) => String(c[0]).includes("analyze-message"));
const sent = () => analyzeCalls().map((c) => JSON.parse(c[1].body));
const actions = () => sent().map((b) => b.action);
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
const answerByAction = () =>
  fetchMock.mockImplementation(async (_u: string, init: { body: string }) =>
    json(200, RESP[JSON.parse(init.body).action])
  );
/* A fetch that stays pending until resolved; rejects on abort like the real one. */
const hangingFetch = () => {
  const calls: {
    signal: AbortSignal;
    action: string;
    body: string;
    resolve: (v: unknown) => void;
  }[] = [];
  fetchMock.mockImplementation(
    (_u: string, init: { signal: AbortSignal; body: string }) =>
      new Promise((resolve, reject) => {
        calls.push({
          signal: init.signal,
          action: JSON.parse(init.body).action,
          body: JSON.parse(init.body).body,
          resolve,
        });
        init.signal.addEventListener("abort", () => reject(new Error("aborted")));
      })
  );
  return calls;
};

beforeEach(() => {
  vi.useFakeTimers();
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

describe("adversarial: toggling, rapid changes, cap mid-session (F-003/AC-62, AC-64, AC-65, AC-68)", () => {
  it("turning the switch off and on again during a pending auto reply sends no second request for that email", async () => {
    setup(makeItem("A", { itemId: "a" }), { ...on("reply"), autoReplyConsent: true });
    const calls = hangingFetch();
    render(<App />);
    tab("Resposta");
    await advance(800);
    expect(calls).toHaveLength(1);
    await turnOn("reply"); // off: aborts the in-flight auto request
    expect(calls[0].signal.aborted).toBe(true);
    await turnOn("reply"); // on again (consent remembered): the pair was already started
    expect(screen.queryByRole("dialog")).toBeNull();
    await advance(5000);
    expect(calls).toHaveLength(1);
  });

  it("(AC-65 b) a request in flight for A is aborted when B arrives; A's late result is never shown; B goes out after its own 800 ms", async () => {
    setup(makeItem("A", { itemId: "a" }), on("reply"));
    const calls = hangingFetch();
    render(<App />);
    tab("Resposta");
    await advance(800);
    expect(calls).toHaveLength(1);
    await change(makeItem("B", { itemId: "b" }));
    expect(calls[0].signal.aborted).toBe(true);
    await act(async () => calls[0].resolve(json(200, RESP.reply))); // late result of A
    expect(screen.queryByText("minha resposta")).toBeNull();
    await advance(799);
    expect(calls).toHaveLength(1);
    await advance(1);
    expect(calls).toHaveLength(2);
    expect(calls[1].body).toBe("B");
    expect(screen.queryByRole("alert")).toBeNull();
  });

  it("rapid item changes (ten within 500 ms) send one request, for the last email only", async () => {
    setup(makeItem("A", { itemId: "a" }), on("reply"));
    render(<App />);
    tab("Resposta");
    for (const b of ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]) {
      await change(makeItem(b, { itemId: b.toLowerCase() }));
      await advance(50);
    }
    await advance(800);
    expect(sent().map((b) => b.body)).toEqual(["K"]);
    await advance(10000);
    expect(analyzeCalls()).toHaveLength(1);
  });

  it("turning the switch off during its 800 ms debounce cancels the request (AC-62)", async () => {
    setup(makeItem("A", { itemId: "a" }), { ...on("reply"), autoReplyConsent: true });
    render(<App />);
    tab("Resposta");
    await advance(400);
    await turnOn("reply"); // off before the debounce ends
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(0);
  });

  it("an email with no itemId: a revisit counts as a new email (per-load key, AC-64)", async () => {
    setup(makeItem("A"), on("reply"));
    render(<App />);
    tab("Resposta");
    await advance(800);
    await change(makeItem("B"));
    await advance(800);
    await change(makeItem("A"));
    await advance(800);
    expect(sent().map((b) => b.body)).toEqual(["A", "B", "A"]);
  });

  it("two switches on, quick Resposta -> Traduzir -> Resposta inside 800 ms: one reply, no translate", async () => {
    setup(makeItem("A", { itemId: "a" }), on("reply", "translate"));
    render(<App />);
    tab("Resposta");
    await advance(300);
    tab("Traduzir");
    await advance(300);
    tab("Resposta");
    await advance(5000);
    expect(actions()).toEqual(["reply"]);
  });

  it("cap reached mid-session: earlier auto results stay, the cap pauses every tab, later emails send nothing", async () => {
    setup(makeItem("A", { itemId: "a" }), on("reply", "translate"));
    let n = 0;
    fetchMock.mockImplementation(async (_u: string, init: { body: string }) => {
      n++;
      if (n === 3) return json(429, { ok: false, error: "cap_reached" });
      return json(200, RESP[JSON.parse(init.body).action]);
    });
    render(<App />);
    tab("Resposta");
    await advance(800); // A ok
    await change(makeItem("B", { itemId: "b" }));
    await advance(800); // B ok
    await change(makeItem("C", { itemId: "c" }));
    await advance(800); // C: cap
    expect(screen.getByRole("alert").textContent).toMatch(
      /Atingiu o limite diário de pedidos de IA/
    );
    expect(screen.getAllByText(/^Modo automático em pausa até amanhã/)).toHaveLength(1);
    tab("Traduzir");
    await advance(5000);
    await change(makeItem("D", { itemId: "d" }));
    tab("Resposta");
    await advance(5000);
    expect(actions()).toEqual(["reply", "reply", "reply"]);
    expect(sw("reply").checked).toBe(true);
    tab("Traduzir");
    expect(sw("translate").checked).toBe(true); // saved switches stay on through the pause
  });

  it("a manual double-click sends exactly one request (AC-69)", async () => {
    setup(makeItem("A", { itemId: "a" }));
    render(<App />);
    tab("Resumo");
    const btn = screen.getByRole("button", { name: "Resumir" });
    await act(async () => {
      fireEvent.click(btn);
      fireEvent.click(btn);
    });
    await advance(2000);
    expect(actions()).toEqual(["summarize"]);
  });
});

describe("adversarial: storage faults and consent values (F-003/AC-60, AC-61, Q-U18)", () => {
  it("roamingSettings.set throws: the choice stays for the session, the notice shows once, nothing is saved", async () => {
    setup(makeItem("corpo"));
    rs.set.mockImplementation(() => {
      throw new Error("boom");
    });
    render(<App />);
    tab("Resposta");
    await turnOn("reply");
    await ligar();
    expect(sw("reply").checked).toBe(true);
    expect(screen.getAllByText(/Não foi possível guardar a preferência/)).toHaveLength(1);
    expect(rs.saveAsync).not.toHaveBeenCalled();
  });

  it("saveAsync throws (not a failed status): same as a failed save, no crash", async () => {
    setup(makeItem("corpo"));
    rs.saveAsync.mockImplementation(() => {
      throw new Error("boom");
    });
    render(<App />);
    tab("Resposta");
    await turnOn("reply");
    await ligar();
    expect(sw("reply").checked).toBe(true);
    expect(screen.getAllByText(/Não foi possível guardar a preferência/)).toHaveLength(1);
  });

  it.each([
    ["the string 'true' in the consent key", { autoReplyConsent: "true" }],
    [
      "the number 1 in the consent key and the string 'true' in the switch key",
      { autoReplyConsent: 1, autoReply: "true" },
    ],
  ])(
    "%s does not count as consent: the dialog opens and nothing is on before Ligar",
    async (_n, store) => {
      setup(makeItem("corpo"), store);
      render(<App />);
      tab("Resposta");
      expect(sw("reply").checked).toBe(false);
      await turnOn("reply");
      expect(screen.getByText(CONSENT_TITLE)).toBeTruthy();
      fireEvent.click(screen.getByRole("button", { name: "Cancelar" }));
      await advance(2000);
      expect(analyzeCalls()).toHaveLength(0);
      expect(sw("reply").checked).toBe(false);
    }
  );

  it("pane restore: a new mount restores the saved switches and consent; the same email is sent again in the new session", async () => {
    const store = { autoReply: true, autoReplyConsent: true };
    setup(makeItem("A", { itemId: "a" }), store);
    render(<App />);
    tab("Resposta");
    expect(sw("reply").checked).toBe(true);
    await advance(800);
    expect(actions()).toEqual(["reply"]);
    cleanup();
    setup(makeItem("A", { itemId: "a" }), store);
    render(<App />);
    tab("Resposta");
    expect(sw("reply").checked).toBe(true);
    expect(screen.queryByRole("dialog")).toBeNull();
    await advance(800);
    expect(actions()).toEqual(["reply", "reply"]);
  });
});

describe("adversarial: request storm and no-reply senders (F-003/AC-63, AC-64, Q-U20)", () => {
  it("all three on, one email, one minute on each tab: exactly one request per tab, no repeats", async () => {
    setup(makeItem("A", { itemId: "a" }), on("summarize", "reply", "translate"));
    render(<App />);
    await advance(800);
    tab("Resposta");
    await advance(60000);
    tab("Traduzir");
    await advance(60000);
    tab("Resumo");
    await advance(60000);
    tab("Lista");
    await advance(60000);
    expect(actions()).toEqual(["summarize", "reply", "translate"]);
  });

  it("a failing (retryable) backend: still one request per tab for the email, no retry loop in three minutes", async () => {
    setup(makeItem("A", { itemId: "a" }), on("summarize", "reply", "translate"));
    fetchMock.mockImplementation(async () =>
      json(503, { ok: false, error: "ai_unavailable", retryable: true })
    );
    render(<App />);
    await advance(800);
    tab("Resposta");
    await advance(60000);
    tab("Traduzir");
    await advance(60000);
    await advance(60000);
    expect(actions()).toEqual(["summarize", "reply", "translate"]);
  });

  it("a no-reply sender on Resposta with the switch on: no request and no alert; the manual button still sends", async () => {
    setup(
      makeItem("A", { from: { displayName: "Sistema", emailAddress: "no-reply@exemplo.pt" } }),
      on("reply")
    );
    render(<App />);
    tab("Resposta");
    await advance(5000);
    expect(analyzeCalls()).toHaveLength(0);
    expect(screen.queryByRole("alert")).toBeNull();
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Sugerir resposta" }));
    });
    expect(actions()).toEqual(["reply"]);
  });
});

describe("adversarial: aborted auto request during the token step (F-003/AC-62)", () => {
  it("a token that resolves OK after the switch was turned off sends nothing", async () => {
    setup(makeItem("A", { itemId: "a" }), { ...on("reply"), autoReplyConsent: true });
    let resolveTok: (v: unknown) => void = () => {};
    (getTokenSilent as any).mockImplementation(() => new Promise((r) => (resolveTok = r)));
    render(<App />);
    tab("Resposta");
    await advance(800);
    expect(getTokenSilent).toHaveBeenCalledTimes(1);
    await turnOn("reply"); // off while the token is still being fetched
    await act(async () => resolveTok({ kind: "token", token: "TOK" }));
    await advance(2000);
    expect(analyzeCalls()).toHaveLength(0);
  });

  // Known defect (AiTools.tsx, after `await getTokenSilent()`): an aborted auto request still shows the
  // sign-in text when the token step fails. Written as it.fails so the suite stays green and documents the
  // defect; when the defect is fixed this test fails and `.fails` must be removed.
  it.fails(
    "DEFECT: a token step that needs sign-in after the switch was turned off shows no sign-in text (AC-62)",
    async () => {
      setup(makeItem("A", { itemId: "a" }), { ...on("reply"), autoReplyConsent: true });
      let resolveTok: (v: unknown) => void = () => {};
      (getTokenSilent as any).mockImplementation(() => new Promise((r) => (resolveTok = r)));
      render(<App />);
      tab("Resposta");
      await advance(800);
      await turnOn("reply"); // off while the token is still being fetched
      await act(async () => resolveTok({ kind: "interaction" }));
      expect(screen.queryByText(SIGN_IN)).toBeNull();
    }
  );
});
