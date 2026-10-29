import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { readFileSync } from "node:fs";
import { join } from "node:path";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

// F-003 AC-40, AC-41, AC-42 (flag off), F-012 AC-43, AC-45, AC-46, AC-50 (pane part). Mocks only.
type Cb = (r: { status: string; value?: string }) => void;
let handler: (() => void) | undefined;
let office: any;
let fetchMock: ReturnType<typeof vi.fn>;
let getAsync: ReturnType<typeof vi.fn>;

const INJECT = "<img src=x onerror=alert(1)> &amp; fim";
const SUMMARY = {
  points: ["a", "b", "c"],
  type: "request",
  priority: "high",
  tone: "neutral",
  deadline: null,
  requestedAction: null,
};
const makeItem = (body: string, o: object = {}) => {
  getAsync = vi.fn((_t: string, cb: Cb) => cb({ status: "succeeded", value: body }));
  return {
    subject: "Assunto",
    from: { displayName: "Ana Silva", emailAddress: "ana@exemplo.pt" },
    dateTimeCreated: new Date(2026, 0, 2),
    body: { getAsync },
    displayReplyFormAsync: vi.fn((_h: string, cb: Cb) => cb({ status: "succeeded" })),
    displayReplyForm: vi.fn(),
    ...o,
  };
};
const setup = (item: unknown) => {
  office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      roamingSettings: { get: vi.fn(() => true), set: vi.fn(), saveAsync: vi.fn() }, // "on" must be ignored: the flag is false
      mailbox: {
        item,
        addHandlerAsync: vi.fn((e: string, h: () => void) => {
          if (e === "itemChanged") handler = h;
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
const tab = (name: string) => fireEvent.click(screen.getByRole("tab", { name }));
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name }));
  });
};
const sent = (i = 0) => JSON.parse(fetchMock.mock.calls[i][1].body);

beforeEach(() => {
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("tabs and context (F-012/AC-43, AC-45)", () => {
  it("tablist with four tabs, Resumo selected, panel labelled by the selected tab", () => {
    setup(makeItem("corpo"));
    render(<App />);
    expect(screen.getByRole("tablist")).toBeTruthy();
    expect(screen.getAllByRole("tab")).toHaveLength(4);
    expect(screen.getByRole("tab", { name: "Resumo" }).getAttribute("aria-selected")).toBe("true");
    tab("Traduzir");
    expect(screen.getByRole("tab", { name: "Traduzir" }).getAttribute("aria-selected")).toBe(
      "true"
    );
    expect(screen.getByRole("tabpanel", { name: "Traduzir" })).toBeTruthy();
  });
  it("reads the body once for the three tool tabs, hides the context on Automático, no request by tab switches", () => {
    setup(makeItem("corpo"));
    render(<App />);
    expect(getAsync).toHaveBeenCalledTimes(1);
    tab("Resposta");
    tab("Traduzir");
    tab("Resumo");
    expect(getAsync).toHaveBeenCalledTimes(1);
    const panel = screen.getByRole("tabpanel", { name: "Resumo" });
    expect(panel.textContent).toContain("Ana Silva");
    tab("Automático");
    expect(panel.hasAttribute("hidden")).toBe(true); // message context and cards are hidden on Automático
    expect(screen.queryByRole("tabpanel", { name: "Resumo" })).toBeNull();
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("shows the privacy notice and the assurance on the three tool tabs", () => {
    setup(makeItem("corpo"));
    render(<App />);
    for (const n of ["Resumo", "Resposta", "Traduzir"]) {
      tab(n);
      expect(screen.getByText(/são enviados à OpenAI/)).toBeTruthy();
      expect(screen.getByText(/Sem envio automático/)).toBeTruthy();
    }
  });
});

describe("Resposta options (F-003/AC-40)", () => {
  it("defaults, change sends nothing, request carries the enums, options survive ItemChanged and tab switches", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "reply", reply: "r" }));
    render(<App />);
    tab("Resposta");
    const lang = screen.getByLabelText("Idioma") as HTMLSelectElement;
    const tone = screen.getByLabelText("Tom") as HTMLSelectElement;
    expect(lang.value).toBe("match");
    expect(tone.value).toBe("none");
    expect(Array.from(lang.options).map((o) => o.text)).toEqual([
      "Igual ao email",
      "Português",
      "English",
    ]);
    expect(Array.from(tone.options).map((o) => o.text)).toEqual([
      "Nenhum",
      "Conciso",
      "Detalhado",
      "Formal",
      "Amigável",
    ]);
    fireEvent.change(lang, { target: { value: "en" } });
    fireEvent.change(tone, { target: { value: "formal" } });
    expect(fetchMock).not.toHaveBeenCalled();
    office.context.mailbox.item = makeItem("outro");
    act(() => handler?.());
    tab("Traduzir");
    tab("Resposta");
    await click("Sugerir resposta");
    expect(sent()).toEqual({
      action: "reply",
      subject: "Assunto",
      fromName: "Ana Silva",
      fromDomain: "exemplo.pt",
      body: "outro",
      language: "en",
      tone: "formal",
    });
  });
  it("sends exactly the selected enum for every language and tone combination", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "reply", reply: "r" }));
    render(<App />);
    tab("Resposta");
    const lang = screen.getByLabelText("Idioma") as HTMLSelectElement;
    const tone = screen.getByLabelText("Tom") as HTMLSelectElement;
    let n = 0;
    for (const l of Array.from(lang.options).map((o) => o.value))
      for (const t of Array.from(tone.options).map((o) => o.value)) {
        fireEvent.change(lang, { target: { value: l } });
        fireEvent.change(tone, { target: { value: t } });
        await click("Sugerir resposta");
        expect(sent(n)).toMatchObject({ language: l, tone: t });
        n++;
      }
    expect(n).toBe(15);
  });
});

describe("Traduzir card (F-003/AC-41, F-012/AC-46)", () => {
  const TR = {
    ok: true,
    action: "translate",
    translation: { subject: "Pedido <b>", body: INJECT },
  };
  it("sends exactly the translate payload with no sender data", async () => {
    setup(makeItem("x".repeat(5000)));
    fetchMock.mockResolvedValue(json(200, TR));
    render(<App />);
    tab("Traduzir");
    expect((screen.getByLabelText("Traduzir para") as HTMLSelectElement).value).toBe("pt");
    fireEvent.change(screen.getByLabelText("Traduzir para"), { target: { value: "en" } });
    expect(fetchMock).not.toHaveBeenCalled();
    await click("Traduzir");
    const body = fetchMock.mock.calls[0][1].body as string;
    expect(Object.keys(sent()).sort()).toEqual(
      ["action", "body", "fromDomain", "fromName", "subject", "targetLanguage"].sort()
    );
    expect(sent()).toMatchObject({
      action: "translate",
      fromName: "",
      fromDomain: "",
      targetLanguage: "en",
    });
    expect(sent().body).toHaveLength(4000);
    expect(body).not.toContain("Ana");
    expect(body).not.toContain("exemplo");
  });
  it("renders subject and body as text, copies subject + blank line + body, has no reply-form button", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(200, TR));
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    const { container } = render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    expect(container.querySelector("img, b")).toBeNull();
    expect(screen.getByText("Assunto: Pedido <b>").getAttribute("dir")).toBe("auto");
    expect(screen.getByText(INJECT).getAttribute("dir")).toBe("auto");
    expect(screen.queryByRole("button", { name: "Abrir resposta para revisão" })).toBeNull();
    await click("Copiar");
    expect(writeText).toHaveBeenCalledWith(`Pedido <b>\n\n${INJECT}`);
    expect(await screen.findByText("Copiado")).toBeTruthy();
  });
  it("omits the subject line when empty and copies the body alone; the error table applies", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValueOnce(
      json(200, { ok: true, action: "translate", translation: { subject: "", body: "só corpo" } })
    );
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    expect(screen.queryByText(/^Assunto:/)).toBeNull();
    await click("Copiar");
    expect(writeText).toHaveBeenCalledWith("só corpo");
    fetchMock.mockResolvedValueOnce(json(502, { ok: false, error: "ai_output_invalid" }));
    await click("Traduzir");
    expect(screen.getByRole("alert").textContent).toContain("A IA devolveu uma resposta inválida");
  });
  it("shows the pending text A traduzir…", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockReturnValue(new Promise(() => undefined));
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    expect(
      (screen.getByRole("button", { name: "A traduzir…" }) as HTMLButtonElement).disabled
    ).toBe(true);
  });
});

describe("ToolCard (F-012/AC-46)", () => {
  it("Copiar and Expandir appear only with a result; Expandir toggles aria-expanded and name; the result is a focusable named region", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(summaryOk());
    render(<App />);
    expect(screen.queryByRole("button", { name: "Copiar" })).toBeNull();
    expect(screen.queryByRole("button", { name: "Expandir" })).toBeNull();
    expect(screen.getByRole("heading", { level: 2, name: "Resumo do email" })).toBeTruthy();
    await click("Resumir");
    const region = screen.getByRole("region", { name: "Resumo do email" });
    expect(region.getAttribute("tabindex")).toBe("0");
    expect(screen.getByRole("button", { name: "Copiar" })).toBeTruthy();
    const expand = screen.getByRole("button", { name: "Expandir" });
    expect(expand.getAttribute("aria-expanded")).toBe("false");
    const collapsedClass = region.className;
    fireEvent.click(expand);
    expect(screen.getByRole("button", { name: "Reduzir" }).getAttribute("aria-expanded")).toBe(
      "true"
    );
    expect(region.className).not.toBe(collapsedClass); // the 240 px max-height class is gone
  });
  it("keeps each card's result when switching tabs", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "reply", reply: "minha resposta" }));
    render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    tab("Traduzir");
    tab("Resposta");
    expect(screen.getByText("minha resposta")).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe("auto mode with the flag off (F-003/AC-42, hidden in production)", () => {
  it("hides the switch and sends no request even with autoSummary = true in roamingSettings", async () => {
    vi.useFakeTimers();
    try {
      setup(makeItem("corpo"));
      render(<App />);
      expect(screen.queryByRole("switch")).toBeNull();
      expect(screen.queryByText("Resumo automático")).toBeNull();
      await act(async () => {
        vi.advanceTimersByTime(5000);
      });
      expect(fetchMock).not.toHaveBeenCalled();
      expect(office.context.roamingSettings.get).not.toHaveBeenCalled();
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("entry page resumo.html (F-012/AC-50)", () => {
  it("summaryEntry sends exactly one summarize request for the first item, none for later items", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(summaryOk());
    render(<App summaryEntry />);
    await act(async () => undefined);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(sent().action).toBe("summarize");
    expect(screen.getByRole("tab", { name: "Resumo" }).getAttribute("aria-selected")).toBe("true");
    office.context.mailbox.item = makeItem("segundo");
    act(() => handler?.());
    await act(async () => undefined);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
  it("sends nothing for an empty body, a null item or a token that needs interaction", async () => {
    setup(makeItem("   "));
    const r = render(<App summaryEntry />);
    await act(async () => undefined);
    r.unmount();
    setup(null);
    const r2 = render(<App summaryEntry />);
    await act(async () => undefined);
    r2.unmount();
    setup(makeItem("corpo"));
    (getTokenSilent as any).mockResolvedValue({ kind: "interaction" });
    render(<App summaryEntry />);
    await act(async () => undefined);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(screen.getByRole("alert").textContent).toContain("Inicie sessão");
  });
  it("without summaryEntry nothing starts", async () => {
    setup(makeItem("corpo"));
    render(<App />);
    await act(async () => undefined);
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("index.tsx derives it from location.pathname only; location.search and hash are never read", async () => {
    for (const f of ["App.tsx", "index.tsx", "AiTools.tsx", "MessagePanel.tsx", "auth.ts"])
      expect(readFileSync(join(process.cwd(), "src/taskpane", f), "utf8")).not.toMatch(
        /location\.(search|hash)/
      );
    window.history.pushState({}, "", "/resumo.html");
    document.body.innerHTML = '<div id="container"></div>';
    setup(makeItem("corpo"));
    let ready: (() => void) | undefined;
    office.onReady = vi.fn((cb: () => void) => (ready = cb));
    fetchMock.mockResolvedValue(summaryOk());
    vi.resetModules();
    await import("./index");
    await act(async () => ready?.());
    await act(async () => undefined);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    window.history.pushState({}, "", "/");
  });
});
