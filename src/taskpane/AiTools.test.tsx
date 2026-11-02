import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { readFileSync } from "node:fs";
import { join } from "node:path";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
const SUMMARY = {
  points: ["<img src=x onerror=alert(1)>", "a &amp; b", "c"],
  type: "request",
  priority: "high",
  tone: "bogus",
  deadline: "sexta",
  requestedAction: null,
};
let handler: (() => void) | undefined;
let office: any;
let fetchMock: ReturnType<typeof vi.fn>;

const makeItem = (body: string, o: object = {}) => ({
  subject: "Assunto",
  from: { displayName: "Ana", emailAddress: "ana@Exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  displayReplyFormAsync: vi.fn((_h: string, cb: Cb) => cb({ status: "succeeded" })),
  displayReplyForm: vi.fn(),
  ...o,
});
const setup = (item: unknown, supported = true) => {
  office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => supported) },
      mailbox: {
        item,
        addHandlerAsync: vi.fn((e: string, h: () => void) => {
          if (e === "itemChanged") handler = h; // the theme hook also registers a handler
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
const ok = (body: unknown) => fetchMock.mockResolvedValue(json(200, body));
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name }));
  });
};
const tab = (name: string) => fireEvent.click(screen.getByRole("tab", { name }));
const btn = (name: string) => screen.getByRole("button", { name }) as HTMLButtonElement;

beforeEach(() => {
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("AiTools layout and no auto-run (F-003/AC-16, AC-17, AC-19)", () => {
  it("shows the Resumo card, notice and assurance; the old cards are gone; no request before a click", () => {
    setup(makeItem("corpo"));
    render(<App />);
    expect(btn("Resumir").disabled).toBe(false);
    expect(screen.queryByRole("button", { name: "Sugerir resposta" })).toBeNull(); // other tab
    expect(screen.getByText(/são enviados à OpenAI/)).toBeTruthy();
    expect(screen.queryByText("Melhorar texto e tom")).toBeNull();
    expect(screen.queryByText("DIMSCALE AI Assistant")).toBeNull();
    expect(screen.getByText("Sem envio automático. O envio é sempre feito por si.")).toBeTruthy();
    expect(screen.queryByText("Resposta rápida")).toBeNull();
    tab("Resposta");
    expect(btn("Sugerir resposta").disabled).toBe(false);
    tab("Traduzir");
    expect(btn("Traduzir").disabled).toBe(false);
    tab("Lista");
    tab("Resumo");
    act(() => handler?.());
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("empty body disables both buttons with the notice; no item or body error shows no buttons", () => {
    setup(makeItem("  "));
    const { unmount } = render(<App />);
    expect(btn("Resumir").disabled).toBe(true);
    tab("Resposta");
    expect(btn("Sugerir resposta").disabled).toBe(true);
    tab("Traduzir");
    expect(btn("Traduzir").disabled).toBe(true);
    expect(
      screen.getAllByText("Esta mensagem não tem texto para analisar.").length
    ).toBeGreaterThan(0);
    unmount();
    setup(null);
    render(<App />);
    expect(screen.queryByRole("button", { name: "Resumir" })).toBeNull();
    cleanup();
    setup(makeItem("x", { body: { getAsync: (_t: string, cb: Cb) => cb({ status: "failed" }) } }));
    render(<App />);
    expect(screen.queryByRole("button", { name: "Resumir" })).toBeNull();
  });
  it("without a token (interaction or error) shows the sign-in text and makes no request", async () => {
    setup(makeItem("corpo"));
    render(<App />);
    (getTokenSilent as any).mockResolvedValue({ kind: "interaction" });
    await click("Resumir");
    expect(screen.getByRole("alert").textContent).toContain(
      "Inicie sessão (em baixo) para usar a IA."
    );
    (getTokenSilent as any).mockResolvedValue({ kind: "error", code: "x" });
    tab("Resposta");
    await click("Sugerir resposta");
    expect(screen.getByRole("alert").textContent).toContain("Inicie sessão");
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe("AiTools request (F-003/AC-17, AC-18)", () => {
  it("sends exactly one POST with the exact body and no address; double-click sends one", async () => {
    setup(makeItem("x".repeat(5000)));
    let release!: (v: unknown) => void;
    fetchMock.mockReturnValue(new Promise((r) => (release = r)));
    render(<App />);
    await act(async () => {
      const b = screen.getByRole("button", { name: "Resumir" });
      fireEvent.click(b);
      fireEvent.click(b);
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(btn("A resumir…").disabled).toBe(true);
    tab("Resposta");
    expect(btn("Sugerir resposta").disabled).toBe(true); // one request at a time per pane
    tab("Resumo");
    const [, init] = fetchMock.mock.calls[0];
    expect(init.headers).toEqual({
      Authorization: "Bearer TOK",
      "Content-Type": "application/json",
    });
    const sent = JSON.parse(init.body);
    expect(Object.keys(sent).sort()).toEqual([
      "action",
      "body",
      "fromDomain",
      "fromName",
      "output",
      "subject",
    ]);
    expect(sent).toMatchObject({
      action: "summarize",
      output: "classified",
      subject: "Assunto",
      fromName: "Ana",
      fromDomain: "exemplo.pt",
    });
    expect(sent.body).toHaveLength(4000);
    expect(init.body).not.toContain("ana@");
    await act(async () => release(json(200, { ok: true, action: "summarize", summary: SUMMARY })));
    expect(btn("Resumir").disabled).toBe(false);
  });
  it("sends an empty subject as an empty string, not the placeholder", async () => {
    setup(makeItem("corpo", { subject: "" }));
    ok({ ok: true, action: "reply", reply: "r" });
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    expect(JSON.parse(fetchMock.mock.calls[0][1].body)).toMatchObject({
      action: "reply",
      subject: "",
      language: "match",
      tone: "none",
    });
  });
});

describe("AiTools results (F-003/AC-20, AC-23, AC-24)", () => {
  it("renders the summary as text with known badges only, deadline only when present", async () => {
    setup(makeItem("corpo"));
    ok({ ok: true, action: "summarize", summary: SUMMARY });
    const { container } = render(<App />);
    await click("Resumir");
    expect(screen.getByText("<img src=x onerror=alert(1)>")).toBeTruthy();
    expect(screen.getByText("a &amp; b")).toBeTruthy();
    expect(container.querySelector("img")).toBeNull();
    expect(container.querySelectorAll("ul li")).toHaveLength(3);
    expect(screen.getByText("Pedido")).toBeTruthy();
    expect(screen.getByText("Prioridade alta")).toBeTruthy();
    expect(screen.queryByText(/^Tom /)).toBeNull();
    expect(screen.getByText("Prazo: sexta")).toBeTruthy();
    expect(screen.queryByText(/^Pedido:/)).toBeNull();
    expect(screen.getByText("Gerado por IA, pode estar errado.")).toBeTruthy();
    expect(container.querySelector("li")?.getAttribute("dir")).toBe("auto");
  });
  it("copies the 3 points as bullet lines, and shows the fallback when the clipboard is missing or rejects", async () => {
    setup(makeItem("corpo"));
    ok({ ok: true, action: "summarize", summary: SUMMARY });
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    render(<App />);
    await click("Resumir");
    await click("Copiar");
    expect(writeText).toHaveBeenCalledWith("• <img src=x onerror=alert(1)>\n• a &amp; b\n• c");
    expect((await screen.findByText("Copiado")).getAttribute("role")).toBe("status");
    writeText.mockRejectedValue(new Error("blocked"));
    await click("Copiar");
    expect(await screen.findByText(/Não foi possível copiar/)).toBeTruthy();
    vi.stubGlobal("navigator", {});
    await click("Copiar");
    expect(screen.getByText(/Não foi possível copiar/)).toBeTruthy();
  });
  it("renders the reply literally, copies it and opens the escaped reply form once (1.9)", async () => {
    const item = makeItem("corpo");
    setup(item);
    const reply = `Olá <b>x</b> & "y"\nAdeus 'z'`;
    ok({ ok: true, action: "reply", reply });
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    const { container } = render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    expect(container.querySelector("b")).toBeNull();
    expect(screen.getByText(/Olá <b>x<\/b>/).getAttribute("dir")).toBe("auto");
    await click("Copiar");
    expect(writeText).toHaveBeenCalledWith(reply);
    await click("Abrir resposta para revisão");
    expect(item.displayReplyFormAsync).toHaveBeenCalledTimes(1);
    expect(item.displayReplyFormAsync.mock.calls[0][0]).toBe(
      "Olá &lt;b&gt;x&lt;/b&gt; &amp; &quot;y&quot;<br>Adeus &#39;z&#39;"
    );
    expect(item.displayReplyForm).not.toHaveBeenCalled();
    expect(screen.queryByText(/Não foi possível abrir/)).toBeNull();
  });
  it("uses displayReplyForm below Mailbox 1.9 and shows the failure text when it throws or the callback fails", async () => {
    const item = makeItem("corpo");
    setup(item, false);
    ok({ ok: true, action: "reply", reply: "r" });
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    await click("Abrir resposta para revisão");
    expect(item.displayReplyForm).toHaveBeenCalledWith("r");
    expect(item.displayReplyFormAsync).not.toHaveBeenCalled();
    item.displayReplyForm.mockImplementation(() => {
      throw new Error("x");
    });
    await click("Abrir resposta para revisão");
    expect(await screen.findByText("Não foi possível abrir a resposta. Use Copiar.")).toBeTruthy();
    cleanup();
    const item2 = makeItem("corpo", {
      displayReplyFormAsync: vi.fn((_h: string, cb: Cb) => cb({ status: "failed" })),
    });
    setup(item2, true);
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    await click("Abrir resposta para revisão");
    expect(await screen.findByText("Não foi possível abrir a resposta. Use Copiar.")).toBeTruthy();
  });
});

describe("AiTools errors (F-003/AC-21)", () => {
  const cases: [number, unknown, string, boolean][] = [
    [
      401,
      { ok: false, error: "unauthorized" },
      "A sessão expirou. Inicie sessão novamente.",
      false,
    ],
    [
      422,
      { ok: false, error: "payload_invalid" },
      "Não foi possível analisar esta mensagem.",
      false,
    ],
    [
      429,
      { ok: false, error: "cap_reached" },
      "Atingiu o limite diário de pedidos de IA. Tente amanhã.",
      false,
    ],
    [503, { ok: false, error: "disabled" }, "A IA está temporariamente desligada.", false],
    [503, { ok: false, error: "ai_not_configured" }, "A IA ainda não está configurada.", false],
    [
      503,
      { ok: false, error: "ai_unavailable", retryable: true },
      "O serviço de IA não respondeu. Tente novamente.",
      true,
    ],
    [
      502,
      { ok: false, error: "ai_output_invalid" },
      "A IA devolveu uma resposta inválida. Tente novamente.",
      true,
    ],
    [
      503,
      { ok: false, error: "ai_unavailable", retryable: false },
      "O serviço de IA recusou o pedido (configuração). Avise o administrador.",
      false,
    ],
    [500, { ok: false, error: "server_error" }, "Erro no serviço DIMSCALE. Tente novamente.", true],
  ];
  it.each(cases)("status %s %j", async (status, body, text, retry) => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(status, body));
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    const alert = screen.getByRole("alert");
    expect(alert.textContent).toContain(text);
    expect(!!screen.queryByRole("button", { name: "Tentar novamente" })).toBe(retry);
    if (retry) {
      fetchMock.mockResolvedValue(json(200, { ok: true, action: "reply", reply: "ok!" }));
      await click("Tentar novamente");
      expect(fetchMock).toHaveBeenCalledTimes(2);
      expect(JSON.parse(fetchMock.mock.calls[1][1].body).action).toBe("reply");
      expect(screen.getByText("ok!")).toBeTruthy();
      expect(screen.queryByRole("alert")).toBeNull();
    }
  });
  it("network error shows the não respondeu text with retry", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockRejectedValue(new Error("net"));
    render(<App />);
    await click("Resumir");
    expect(screen.getByRole("alert").textContent).toContain("O serviço de IA não respondeu.");
    expect(screen.getByRole("button", { name: "Tentar novamente" })).toBeTruthy();
  });
});

describe("AiTools honest AI errors on every action (F-003 amendment (g))", () => {
  const REJECTED = "O serviço de IA recusou o pedido (configuração). Avise o administrador.";
  // [tab to open, button that runs the action]
  const actions: [string, string][] = [
    ["Resumo", "Resumir"],
    ["Resposta", "Sugerir resposta"],
    ["Traduzir", "Traduzir"],
  ];
  const err = (status: number, body: object) => json(status, { ok: false, ...body });
  const open = (tabName: string) => {
    setup(makeItem("corpo"));
    render(<App />);
    tab(tabName);
  };
  const spies: ReturnType<typeof vi.spyOn>[] = [];
  const watchConsole = () =>
    (["log", "info", "warn", "error", "debug"] as const).forEach((m) =>
      spies.push(vi.spyOn(console, m).mockImplementation(() => undefined))
    );
  afterEach(() => {
    spies.forEach((sp) => sp.mockRestore());
    spies.length = 0;
  });
  const noConsole = () => spies.forEach((sp) => expect(sp).not.toHaveBeenCalled());

  it.each(actions)("%s: a rejected request shows the new text and the code, with no retry", async (t, b) => {
    watchConsole();
    fetchMock.mockResolvedValue(
      err(503, { error: "ai_unavailable", retryable: false, why: "http_401" })
    );
    open(t);
    await click(b);
    const text = screen.getByRole("alert").textContent ?? "";
    expect(text).toContain(REJECTED);
    expect(text).toContain("(código: http_401)");
    expect(text).not.toContain("inválida");
    expect(screen.queryByRole("button", { name: "Tentar novamente" })).toBeNull();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    noConsole();
  });
  it.each(actions)("%s: a rejected request without a why shows the text and no code", async (t, b) => {
    fetchMock.mockResolvedValue(err(503, { error: "ai_unavailable", retryable: false }));
    open(t);
    await click(b);
    expect(screen.getByRole("alert").textContent).toContain(REJECTED);
    expect(screen.getByRole("alert").textContent).not.toContain("código");
    expect(screen.queryByRole("button", { name: "Tentar novamente" })).toBeNull();
  });
  it.each(actions)("%s: unavailable and invalid keep a retry button and show their code", async (t, b) => {
    open(t);
    for (const [status, body, code] of [
      [503, { error: "ai_unavailable", retryable: true, why: "http_429" }, "http_429"],
      [503, { error: "ai_unavailable", retryable: true, why: "timeout" }, "timeout"],
      [503, { error: "ai_unavailable", retryable: true, why: "network" }, "network"],
      [502, { error: "ai_output_invalid", why: "badbody" }, "badbody"],
    ] as [number, object, string][]) {
      fetchMock.mockResolvedValue(err(status, body));
      await click(b);
      expect(screen.getByRole("alert").textContent).toContain(`(código: ${code})`);
      expect(screen.getByRole("button", { name: "Tentar novamente" })).toBeTruthy();
    }
  });
  it.each(actions)("%s: an unknown why is never shown", async (t, b) => {
    watchConsole();
    open(t);
    const bad = ["<b>evil</b>", "Invalid API key sk-123", "http_4", "http_4011", "timeout", 42, {}];
    for (const why of bad) {
      fetchMock.mockResolvedValue(err(503, { error: "ai_unavailable", retryable: false, why }));
      await click(b);
      expect(screen.getByRole("alert").textContent).toContain(REJECTED);
      expect(screen.getByRole("alert").textContent).not.toContain("código");
      expect(screen.queryByText(/evil|sk-123/)).toBeNull();
    }
    for (const why of ["<b>evil</b>", "http_401", "Schema", 7]) {
      fetchMock.mockResolvedValue(err(502, { error: "ai_output_invalid", why }));
      await click(b);
      expect(screen.getByRole("alert").textContent).not.toContain("código");
      expect(screen.queryByText(/evil/)).toBeNull();
    }
    noConsole();
  });
});

describe("AiTools item change (F-003/AC-22)", () => {
  it("clears results and discards a late response for the previous item", async () => {
    setup(makeItem("one"));
    let release!: (v: unknown) => void;
    fetchMock.mockReturnValueOnce(new Promise((r) => (release = r)));
    render(<App />);
    await click("Resumir");
    office.context.mailbox.item = makeItem("two", { subject: "Outro" });
    act(() => handler?.());
    await act(async () => release(json(200, { ok: true, action: "summarize", summary: SUMMARY })));
    expect(screen.queryByText("a &amp; b")).toBeNull();
    expect(btn("Resumir").disabled).toBe(false);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    ok({ ok: true, action: "summarize", summary: SUMMARY });
    await click("Resumir");
    expect(screen.getByText("a &amp; b")).toBeTruthy();
    office.context.mailbox.item = makeItem("three");
    act(() => handler?.());
    await waitFor(() => expect(screen.queryByText("a &amp; b")).toBeNull());
  });
});

describe("source guards (F-003/AC-24, AC-25)", () => {
  const src = ["AiTools.tsx", "analyzeClient.ts", "replyForm.ts", "messageView.ts"]
    .map((f) => readFileSync(join(process.cwd(), "src/taskpane", f), "utf8"))
    .join("\n");
  it("has no console, storage, HTML injection, reply-all, forward or send API", () => {
    expect(src).not.toMatch(/console\.|localStorage|sessionStorage|cookie|indexedDB/);
    expect(src).not.toMatch(/dangerouslySetInnerHTML|innerHTML|insertAdjacentHTML/);
    expect(src).not.toMatch(/displayReplyAll|displayNewMessage|sendAsync|Mail\.Send|\/send/);
  });
});
