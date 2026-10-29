import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

// QA (independent): pane state, pending, item change, timeouts, clipboard (F-003/AC-21, AC-22, AC-23; F-012/AC-35). Mocks only.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent, signIn } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
let handlers: Set<() => void>;
let office: any;
let fetchMock: ReturnType<typeof vi.fn>;

const SUMMARY = {
  points: ["um", "dois", "tres"],
  type: "request",
  priority: "high",
  tone: "neutral",
  deadline: null,
  requestedAction: null,
};
const makeItem = (body: string, o: object = {}) => ({
  subject: "Assunto",
  from: { displayName: "Ana", emailAddress: "ana@exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  displayReplyFormAsync: vi.fn((_h: string, cb: Cb) => cb({ status: "succeeded" })),
  displayReplyForm: vi.fn(),
  ...o,
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
  };
  (globalThis as any).Office = office;
};
const json = (status: number, body: unknown) => ({
  ok: status < 300,
  status,
  json: async () => body,
});
const okSummary = () => json(200, { ok: true, action: "summarize", summary: SUMMARY });
const okReply = (reply = "resposta") => json(200, { ok: true, action: "reply", reply });
const okTranslate = (body = "traducao") =>
  json(200, { ok: true, action: "translate", translation: { subject: "", body } });
const change = async (item: unknown) => {
  office.context.mailbox.item = item;
  await act(async () => handlers.forEach((h) => h()));
};
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name }));
  });
};
const tab = (name: string) => fireEvent.click(screen.getByRole("tab", { name }));
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
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.clearAllMocks();
});

describe("item change clears every card (F-003/AC-22, F-012/AC-35)", () => {
  it("a late REPLY for the previous item is discarded; the new item's reply card is empty", async () => {
    setup(makeItem("A"));
    let release!: (v: unknown) => void;
    fetchMock.mockReturnValueOnce(new Promise((r) => (release = r)));
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    await change(makeItem("B", { subject: "B" }));
    await act(async () => release(okReply("RESPOSTA-DE-A")));
    expect(screen.queryByText("RESPOSTA-DE-A")).toBeNull();
    expect(screen.queryByRole("button", { name: "Abrir resposta para revisão" })).toBeNull();
    expect(screen.getByRole("button", { name: "Sugerir resposta" })).toBeTruthy();
  });

  it("a late TRANSLATION for the previous item is discarded", async () => {
    setup(makeItem("A"));
    let release!: (v: unknown) => void;
    fetchMock.mockReturnValueOnce(new Promise((r) => (release = r)));
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    await change(makeItem("B"));
    await act(async () => release(okTranslate("TRADUCAO-DE-A")));
    expect(screen.queryByText("TRADUCAO-DE-A")).toBeNull();
    expect(screen.getByRole("button", { name: "Traduzir" })).toBeTruthy();
  });

  it("an error shown for A is cleared when B is selected", async () => {
    setup(makeItem("A"));
    fetchMock.mockResolvedValue(json(500, { ok: false, error: "server_error" }));
    render(<App />);
    await click("Resumir");
    expect(screen.getByRole("alert")).toBeTruthy();
    await change(makeItem("B"));
    expect(screen.queryByRole("alert")).toBeNull();
  });

  it("a result that arrives while another tab is selected is kept and shown when the user returns", async () => {
    setup(makeItem("A"));
    let release!: (v: unknown) => void;
    fetchMock.mockReturnValueOnce(new Promise((r) => (release = r)));
    render(<App />);
    await click("Resumir");
    tab("Resposta");
    await act(async () => release(okSummary()));
    expect(screen.queryByText("dois")).toBeTruthy(); // in the DOM (hidden), not visible
    tab("Resumo");
    expect(screen.getByText("dois")).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe("pending state (F-003/AC-17, AC-22)", () => {
  it("double-click on Traduzir and on Sugerir resposta sends one request each", async () => {
    setup(makeItem("A"));
    fetchMock.mockReturnValue(new Promise(() => undefined));
    render(<App />);
    tab("Traduzir");
    await act(async () => {
      const b = screen.getByRole("button", { name: "Traduzir" });
      fireEvent.click(b);
      fireEvent.click(b);
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    tab("Resposta");
    await act(async () => {
      const b = screen.getByRole("button", { name: "Sugerir resposta" });
      fireEvent.click(b);
      fireEvent.click(b);
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("while a request is pending, every card's primary button is disabled and the pending button is aria-busy", async () => {
    setup(makeItem("A"));
    fetchMock.mockReturnValue(new Promise(() => undefined));
    render(<App />);
    await click("Resumir");
    const busy = screen.getByRole("button", { name: "A resumir…" }) as HTMLButtonElement;
    expect(busy.disabled).toBe(true);
    expect(busy.getAttribute("aria-busy")).toBe("true");
    tab("Resposta");
    expect(
      (screen.getByRole("button", { name: "Sugerir resposta" }) as HTMLButtonElement).disabled
    ).toBe(true);
    tab("Traduzir");
    expect((screen.getByRole("button", { name: "Traduzir" }) as HTMLButtonElement).disabled).toBe(
      true
    );
  });

  it("the pending request is not aborted by a tab switch and its result lands on its card", async () => {
    setup(makeItem("A"));
    const calls = hangingFetch();
    render(<App />);
    await click("Resumir");
    tab("Traduzir");
    expect(calls[0].signal.aborted).toBe(false);
    await act(async () => calls[0].resolve(okSummary()));
    tab("Resumo");
    expect(screen.getByText("tres")).toBeTruthy();
  });

  it("signIn (the popup) is never called by any AI action, including errors and sign-in needed", async () => {
    setup(makeItem("A"));
    (getTokenSilent as any).mockResolvedValue({ kind: "interaction" });
    render(<App />);
    await click("Resumir");
    tab("Resposta");
    await click("Sugerir resposta");
    tab("Traduzir");
    await click("Traduzir");
    expect(signIn).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe("timeouts and malformed responses (F-003/AC-21)", () => {
  it("no answer for 40 s: the 'não respondeu' text with Tentar novamente, one request, nothing before 40 s", async () => {
    vi.useFakeTimers();
    setup(makeItem("A"));
    const calls = hangingFetch();
    render(<App />);
    await act(async () => {
      fireEvent.click(screen.getByRole("button", { name: "Resumir" }));
    });
    await act(async () => {
      await vi.advanceTimersByTimeAsync(39999);
    });
    expect(screen.queryByRole("alert")).toBeNull();
    await act(async () => {
      await vi.advanceTimersByTimeAsync(1);
    });
    expect(calls[0].signal.aborted).toBe(true);
    expect(screen.getByRole("alert").textContent).toContain(
      "O serviço de IA não respondeu. Tente novamente."
    );
    expect(screen.getByRole("button", { name: "Tentar novamente" })).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it.each([
    ["200 with malformed JSON", 200],
    ["502 with malformed JSON", 502],
    ["503 with malformed JSON", 503],
  ])("%s shows the generic error and offers retry", async (_name, status) => {
    setup(makeItem("A"));
    fetchMock.mockResolvedValue({
      ok: status < 300,
      status,
      json: async () => {
        throw new SyntaxError("Unexpected token < in JSON");
      },
    });
    render(<App />);
    await click("Resumir");
    expect(screen.getByRole("alert").textContent).toContain(
      "Erro no serviço DIMSCALE. Tente novamente."
    );
    expect(screen.getByRole("button", { name: "Tentar novamente" })).toBeTruthy();
    // the SyntaxError text is never shown to the user
    expect(document.body.textContent).not.toContain("Unexpected token");
  });

  it("a 200 summary with a non-string point shows the generic error in the Resumo card", async () => {
    setup(makeItem("A"));
    fetchMock.mockResolvedValue(
      json(200, { ok: true, action: "summarize", summary: { ...SUMMARY, points: ["a", 7, "c"] } })
    );
    render(<App />);
    await click("Resumir");
    expect(screen.getByRole("alert").textContent).toContain("Erro no serviço DIMSCALE.");
    expect(screen.queryByRole("list")).toBeNull();
  });
});

describe("clipboard failure paths (F-003/AC-23, F-012/AC-46)", () => {
  it("reply copy: a rejected promise shows the fallback; a later success shows Copiado", async () => {
    setup(makeItem("A"));
    fetchMock.mockResolvedValue(okReply("texto"));
    const writeText = vi
      .fn()
      .mockRejectedValueOnce(new Error("blocked"))
      .mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    await click("Copiar");
    expect(
      (
        await screen.findByText("Não foi possível copiar. Selecione o texto e use Ctrl+C.")
      ).getAttribute("role")
    ).toBe("status");
    await click("Copiar");
    expect((await screen.findByText("Copiado")).getAttribute("role")).toBe("status");
    expect(writeText).toHaveBeenLastCalledWith("texto");
  });

  it("translate copy: a SYNCHRONOUS throw from writeText is caught and shows the fallback", async () => {
    setup(makeItem("A"));
    fetchMock.mockResolvedValue(okTranslate("corpo"));
    vi.stubGlobal("navigator", {
      clipboard: {
        writeText: () => {
          throw new Error("sync");
        },
      },
    });
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    await click("Copiar");
    expect((await screen.findByText(/Não foi possível copiar/)).getAttribute("role")).toBe(
      "status"
    );
  });

  it("no navigator.clipboard at all: the fallback text, and the result stays on screen", async () => {
    setup(makeItem("A"));
    fetchMock.mockResolvedValue(okSummary());
    vi.stubGlobal("navigator", {});
    render(<App />);
    await click("Resumir");
    await click("Copiar");
    expect(screen.getByText(/Não foi possível copiar/)).toBeTruthy();
    expect(screen.getByText("dois")).toBeTruthy();
  });
});

describe("at most one AI request in flight across an item change (F-003/AC-17, AC-44; fixed in 41e10e8)", () => {
  it("a manual request for A still in flight is aborted when B is selected; a click on B then sends one live request", async () => {
    setup(makeItem("A"));
    const calls = hangingFetch();
    render(<App />);
    await click("Resumir");
    expect(calls).toHaveLength(1);
    await change(makeItem("B"));
    expect(calls[0].signal.aborted).toBe(true);
    await click("Resumir");
    expect(calls).toHaveLength(2);
    expect(calls.filter((c) => !c.signal.aborted)).toHaveLength(1);
  });

  it("a fetch that ignores the abort and answers late does not reach B's card", async () => {
    setup(makeItem("A"));
    let release!: (v: unknown) => void;
    fetchMock.mockReturnValueOnce(new Promise((r) => (release = r))); // no abort listener: a webview that ignores abort
    render(<App />);
    await click("Resumir");
    await change(makeItem("B"));
    await act(async () => release(okSummary()));
    expect(screen.queryByText("dois")).toBeNull();
    expect(screen.getByRole("button", { name: "Resumir" })).toBeTruthy();
  });
});

describe("a token that arrives after the user left the email (F-003/AC-22, privacy)", () => {
  it("manual click on A, token still pending, the user selects B, the token arrives: no request is sent for A", async () => {
    setup(makeItem("A"));
    let resolveToken!: (v: unknown) => void;
    (getTokenSilent as any).mockImplementationOnce(() => new Promise((r) => (resolveToken = r)));
    render(<App />);
    await click("Resumir");
    await change(makeItem("B"));
    await act(async () => resolveToken({ kind: "token", token: "TOK" }));
    await act(async () => undefined);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(screen.getByRole("button", { name: "Resumir" })).toBeTruthy();
  });
});
