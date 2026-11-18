import * as React from "react";
import { readdirSync, readFileSync } from "fs";
import { join } from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";
import { AI_TEXT_LIMIT } from "./messageView";
import { strings } from "./strings";

// F-010 compose mode. Mocked Office.js compose item; every recipient/sender member throws if touched (AC-6).
type R = { status: string; value?: any };
let handler: (() => void) | undefined;
let fetchMock: ReturnType<typeof vi.fn>;
let item: any;
let mailbox: any;

const ok = (value?: any): R => ({ status: "succeeded", value });
const makeDraft = (
  o: { subject?: string; body?: string; type?: string; prepend?: R | "throw" } = {}
) => {
  const d: any = {
    subject: { getAsync: vi.fn((cb: (r: R) => void) => cb(ok(o.subject ?? "RE: Proposta"))) },
    body: {
      getAsync: vi.fn((_t: string, cb: (r: R) => void) => cb(ok(o.body ?? "Olá\n> citado"))),
      getTypeAsync: vi.fn((cb: (r: R) => void) => cb(ok(o.type ?? "html"))),
      prependAsync: vi.fn((_d: string, _o: unknown, cb: (r: R) => void) => {
        if (o.prepend === "throw") throw new Error("boom with content");
        cb(o.prepend ?? ok());
      }),
      setAsync: vi.fn(),
      setSelectedDataAsync: vi.fn(),
    },
    displayReplyForm: vi.fn(),
    displayReplyFormAsync: vi.fn(),
  };
  for (const k of ["to", "cc", "bcc", "from", "sender", "recipients", "itemId"])
    Object.defineProperty(d, k, {
      get() {
        throw new Error(`compose must not read ${k}`);
      },
    });
  return d;
};
const setOffice = (it: any) => {
  item = it;
  mailbox = {
    item,
    addHandlerAsync: vi.fn((e: string, h: () => void) => {
      if (e === "itemChanged") handler = h;
    }),
    removeHandlerAsync: vi.fn(),
  };
  (globalThis as any).Office = {
    CoercionType: { Text: "text", Html: "html" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: { requirements: { isSetSupported: vi.fn(() => true) }, mailbox },
  };
};
const res = (reply: string) => ({
  ok: true,
  status: 200,
  json: async () => ({ ok: true, action: "reply", reply }),
});
const btn = (name: string) => screen.getByRole("button", { name }) as HTMLButtonElement;
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(btn(name));
  });
};
const body = (i: number) => JSON.parse(fetchMock.mock.calls[i][1].body);

beforeEach(() => {
  fetchMock = vi.fn().mockResolvedValue(res("Bom dia Ana"));
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
  handler = undefined;
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("mode detection and view (F-010/AC-4, AC-5)", () => {
  it("shows only the reply card for a compose item, without tabs, auto switch or Lista", () => {
    setOffice(makeDraft());
    render(<App />);
    expect(screen.getByRole("heading", { name: strings.composeTitle })).toBeTruthy();
    expect(screen.queryByRole("tablist")).toBeNull();
    for (const t of ["Resumo", "Traduzir", "Lista"]) expect(screen.queryByText(t)).toBeNull();
    expect(screen.queryByRole("switch")).toBeNull();
    expect(screen.getByText(strings.composePrivacyNotice)).toBeTruthy();
    expect(screen.getByLabelText("Idioma")).toBeTruthy();
    const tones = Array.from(
      (screen.getByLabelText("Tom") as HTMLSelectElement).options,
      (o) => o.value
    );
    expect(tones).toEqual([
      "none",
      "professional",
      "concise",
      "detailed",
      "formal",
      "friendly",
      "carnegie",
    ]);
    expect(btn("Sugerir resposta")).toBeTruthy();
    expect(screen.queryByRole("button", { name: strings.composeInsert })).toBeNull(); // no result yet
  });

  it("keeps the read pane for a string-subject item and the no-item state for null", () => {
    setOffice({
      subject: "Olá",
      from: { displayName: "Ana", emailAddress: "a@x.pt" },
      dateTimeCreated: new Date(2026, 0, 2),
      body: { getAsync: (_t: string, cb: (r: R) => void) => cb(ok("corpo")) },
    });
    render(<App />);
    expect(screen.getByRole("tablist")).toBeTruthy();
    expect(screen.queryByRole("heading", { name: strings.composeTitle })).toBeNull();
    mailbox.item = null;
    act(() => handler?.());
    expect(screen.getByText(strings.noItem)).toBeTruthy();
  });

  it("re-detects on ItemChanged: read -> compose -> read, handler removed on unmount", () => {
    setOffice({
      subject: "Olá",
      from: { displayName: "Ana", emailAddress: "a@x.pt" },
      dateTimeCreated: new Date(2026, 0, 2),
      body: { getAsync: (_t: string, cb: (r: R) => void) => cb(ok("corpo")) },
    });
    const { unmount } = render(<App />);
    mailbox.item = makeDraft();
    act(() => handler?.());
    expect(screen.queryByRole("tablist")).toBeNull();
    expect(screen.getByRole("button", { name: "Sugerir resposta" })).toBeTruthy();
    mailbox.item = {
      subject: "Outro",
      body: { getAsync: (_t: string, cb: (r: R) => void) => cb(ok("c")) },
    };
    act(() => handler?.());
    expect(screen.getByRole("tablist")).toBeTruthy();
    unmount();
    expect(mailbox.removeHandlerAsync).toHaveBeenCalled();
  });

  it("Idioma and Tom persist for the pane session across ItemChanged", async () => {
    setOffice(makeDraft());
    render(<App />);
    fireEvent.change(screen.getByLabelText("Tom"), { target: { value: "formal" } });
    fireEvent.change(screen.getByLabelText("Idioma"), { target: { value: "en" } });
    mailbox.item = makeDraft();
    act(() => handler?.());
    expect((screen.getByLabelText("Tom") as HTMLSelectElement).value).toBe("formal");
    expect((screen.getByLabelText("Idioma") as HTMLSelectElement).value).toBe("en");
    await click("Sugerir resposta");
    expect(body(0)).toMatchObject({ language: "en", tone: "formal" });
  });

  it("makes no request on render, ItemChanged, option changes or sign-in (AC-5)", () => {
    setOffice(makeDraft());
    render(<App />);
    fireEvent.change(screen.getByLabelText("Tom"), { target: { value: "formal" } });
    fireEvent.change(screen.getByLabelText("Idioma"), { target: { value: "en" } });
    mailbox.item = makeDraft();
    act(() => handler?.());
    expect(fetchMock).not.toHaveBeenCalled();
    expect(getTokenSilent).not.toHaveBeenCalled();
    expect(item.body.prependAsync).not.toHaveBeenCalled();
  });
});

describe("request (F-010/AC-5, AC-6, AC-7)", () => {
  it("one click = one POST with the exact payload, headers and no recipient access; double click = one", async () => {
    setOffice(makeDraft({ subject: "RE: Proposta", body: "  Olá\n> citado  " }));
    render(<App />);
    fireEvent.change(screen.getByLabelText("Tom"), { target: { value: "carnegie" } });
    await act(async () => {
      fireEvent.click(btn("Sugerir resposta"));
      fireEvent.click(btn("Sugerir resposta"));
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(body(0)).toEqual({
      action: "reply",
      subject: "RE: Proposta",
      fromName: "",
      fromDomain: "",
      body: "Olá\n> citado",
      language: "match",
      tone: "carnegie",
    });
    expect(item.body.getAsync.mock.calls[0][0]).toBe("text");
    const init = fetchMock.mock.calls[0][1];
    expect(init.headers).toEqual({
      Authorization: "Bearer TOK",
      "Content-Type": "application/json",
    });
    expect(screen.getByText("Bom dia Ana")).toBeTruthy();
  });

  it("cuts subject to 200 and body to AI_TEXT_LIMIT without a lone surrogate", async () => {
    expect(AI_TEXT_LIMIT).toBe(12000);
    setOffice(makeDraft({ subject: "s".repeat(300), body: "a".repeat(AI_TEXT_LIMIT - 1) + "😀xyz" }));
    render(<App />);
    await click("Sugerir resposta");
    expect(body(0).subject).toHaveLength(200);
    expect(body(0).body).toHaveLength(AI_TEXT_LIMIT - 1); // the emoji's high surrogate was dropped
  });

  it("empty body: note and no request; the button works again once the draft has text", async () => {
    setOffice(makeDraft({ body: " \n " }));
    render(<App />);
    await click("Sugerir resposta");
    expect(screen.getByRole("alert").textContent).toContain(strings.composeEmpty);
    expect(fetchMock).not.toHaveBeenCalled();
    item.body.getAsync.mockImplementation((_t: string, cb: (r: R) => void) => cb(ok("agora tem")));
    await click("Sugerir resposta");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("a failing getAsync shows the read error, sends nothing and logs nothing", async () => {
    const log = vi.spyOn(console, "error");
    const warn = vi.spyOn(console, "log");
    setOffice(makeDraft());
    item.subject.getAsync.mockImplementation((cb: (r: R) => void) => cb({ status: "failed" }));
    render(<App />);
    await click("Sugerir resposta");
    expect(screen.getByRole("alert").textContent).toContain(strings.composeReadFailed);
    item.subject.getAsync.mockImplementation(() => {
      throw new Error("secret subject");
    });
    await click("Sugerir resposta");
    expect(screen.getByRole("alert").textContent).toContain(strings.composeReadFailed);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(log).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
  });

  it("sign-in needed shows the message and sends nothing", async () => {
    (getTokenSilent as any).mockResolvedValue({ kind: "interaction" });
    setOffice(makeDraft());
    render(<App />);
    await click("Sugerir resposta");
    expect(screen.getByRole("alert").textContent).toContain(strings.aiNeedSignIn);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("shows the F-003 error texts and Tentar novamente only where retryable", async () => {
    setOffice(makeDraft());
    render(<App />);
    fetchMock.mockResolvedValue({
      ok: false,
      status: 503,
      json: async () => ({ error: "ai_unavailable", retryable: false }),
    });
    await click("Sugerir resposta");
    expect(screen.getByRole("alert").textContent).toContain(strings.aiErrRejected);
    expect(screen.queryByRole("button", { name: strings.retry })).toBeNull();
    fetchMock.mockResolvedValue({
      ok: false,
      status: 503,
      json: async () => ({ error: "ai_unavailable", retryable: true }),
    });
    await click("Sugerir resposta");
    expect(screen.getByRole("alert").textContent).toContain(strings.aiErrUnavailable);
    expect(screen.getByRole("button", { name: strings.retry })).toBeTruthy();
  });

  it("renders the suggestion as text (AC-8)", async () => {
    const evil = '<img src=x onerror=alert(1)> &amp;';
    fetchMock.mockResolvedValue(res(evil));
    setOffice(makeDraft());
    const { container } = render(<App />);
    await click("Sugerir resposta");
    expect(screen.getByText(evil)).toBeTruthy();
    expect(container.querySelector("img")).toBeNull();
    expect(screen.getByText(strings.aiAdvisory)).toBeTruthy();
  });

  it("ItemChanged aborts the in-flight request and a late response is never shown", async () => {
    let resolve!: (v: unknown) => void;
    let signal!: AbortSignal;
    fetchMock.mockImplementation((_u: string, init: { signal: AbortSignal }) => {
      signal = init.signal;
      return new Promise((r) => (resolve = r));
    });
    setOffice(makeDraft());
    render(<App />);
    await click("Sugerir resposta");
    mailbox.item = makeDraft();
    act(() => handler?.());
    expect(signal.aborted).toBe(true);
    await act(async () => resolve(res("DO RASCUNHO ANTERIOR")));
    expect(screen.queryByText("DO RASCUNHO ANTERIOR")).toBeNull();
    expect(btn("Sugerir resposta").disabled).toBe(false);
  });
});

describe("refine, previous version, copy (F-010/AC-9, AC-10)", () => {
  it("refine sends the snapshot of the first click, not the edited draft; a new click reads again", async () => {
    setOffice(makeDraft({ subject: "RE: A", body: "texto original" }));
    render(<App />);
    await click("Sugerir resposta");
    await click(strings.composeInsert); // the draft now has the inserted text on top
    item.body.getAsync.mockImplementation((_t: string, cb: (r: R) => void) =>
      cb(ok("Bom dia Ana\n\ntexto original"))
    );
    item.subject.getAsync.mockImplementation((cb: (r: R) => void) => cb(ok("RE: B")));
    fetchMock.mockResolvedValue(res("v1"));
    fireEvent.change(screen.getByPlaceholderText("Ex.: mais curto, pede o prazo"), {
      target: { value: "  mais curto " },
    });
    await click("Enviar alterações");
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(body(1)).toEqual({
      action: "reply",
      subject: "RE: A",
      fromName: "",
      fromDomain: "",
      body: "texto original",
      language: "match",
      tone: "none",
      instruction: "mais curto",
      draft: "Bom dia Ana",
    });
    expect(btn(strings.composeInsert).disabled).toBe(false); // a revision can be inserted
    await click("Versão anterior");
    expect(screen.getByText("Bom dia Ana")).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(2); // no request for the previous version
    await click("Sugerir resposta");
    expect(body(2)).toMatchObject({ subject: "RE: B", body: "Bom dia Ana\n\ntexto original" });
    expect(body(2).instruction).toBeUndefined();
  });

  it("a failed new Sugerir does not replace the refine snapshot", async () => {
    setOffice(makeDraft({ subject: "RE: A", body: "texto original" }));
    render(<App />);
    await click("Sugerir resposta");
    item.body.getAsync.mockImplementation((_t: string, cb: (r: R) => void) => cb(ok("texto novo")));
    fetchMock.mockResolvedValue({
      ok: false,
      status: 503,
      json: async () => ({ error: "ai_unavailable", retryable: true }),
    });
    await click("Sugerir resposta");
    expect(screen.getByRole("alert")).toBeTruthy();
    fetchMock.mockResolvedValue(res("v1"));
    fireEvent.change(screen.getByPlaceholderText("Ex.: mais curto, pede o prazo"), {
      target: { value: "mais curto" },
    });
    await click("Enviar alterações");
    expect(body(2)).toMatchObject({ body: "texto original", instruction: "mais curto" });
  });

  it("Copiar writes the text; a missing clipboard shows the fallback", async () => {
    setOffice(makeDraft());
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    render(<App />);
    await click("Sugerir resposta");
    await click("Copiar");
    expect(writeText).toHaveBeenCalledWith("Bom dia Ana");
    expect(screen.getByText(strings.aiCopied).getAttribute("role")).toBe("status");
    writeText.mockRejectedValue(new Error("no"));
    await click("Copiar");
    expect(screen.getByText(strings.aiCopyFailed)).toBeTruthy();
  });
});

describe("insert (F-010/AC-11, AC-12)", () => {
  const special = `a & b <c> "d" 'e'\r\nlinha 2\n<img src=x onerror=alert(1)>`;

  it("HTML body: one prependAsync with escaped text, <br> and a trailing <br><br>", async () => {
    fetchMock.mockResolvedValue(res(special));
    setOffice(makeDraft({ type: "html" }));
    render(<App />);
    await click("Sugerir resposta");
    expect(item.body.prependAsync).not.toHaveBeenCalled(); // result arrival does not insert
    await act(async () => {
      fireEvent.click(btn(strings.composeInsert));
      fireEvent.click(btn(strings.composeInsert));
    });
    expect(item.body.getTypeAsync).toHaveBeenCalledTimes(1);
    expect(item.body.prependAsync).toHaveBeenCalledTimes(1);
    const [data, opts] = item.body.prependAsync.mock.calls[0];
    expect(data).toBe(
      "a &amp; b &lt;c&gt; &quot;d&quot; &#39;e&#39;<br>linha 2<br>&lt;img src=x onerror=alert(1)&gt;<br><br>"
    );
    expect(opts).toEqual({ coercionType: "html" });
    expect(screen.getByText(strings.composeInserted)).toBeTruthy();
    expect(btn(strings.composeInsert).disabled).toBe(true); // no double insert
  });

  it("insert help is linked to the button; status regions are mounted before their text appears", async () => {
    setOffice(makeDraft());
    vi.stubGlobal("navigator", { clipboard: { writeText: vi.fn().mockResolvedValue(undefined) } });
    render(<App />);
    await click("Sugerir resposta");
    const help = screen.getByText(strings.composeInsertHelp);
    expect(btn(strings.composeInsert).getAttribute("aria-describedby")).toBe(help.id);
    const empty = () => screen.getAllByRole("status").filter((e) => e.textContent === "" && e.tagName === "SPAN");
    expect(empty().length).toBeGreaterThanOrEqual(3); // inserted, copied, refine regions exist with no text
    const before = empty();
    await click(strings.composeInsert);
    await click("Copiar");
    const inserted = screen.getByText(strings.composeInserted);
    const copied = screen.getByText(strings.aiCopied);
    expect(before).toContain(inserted); // same nodes, text changed
    expect(before).toContain(copied);
  });

  it("text body: the plain text followed by a blank line", async () => {
    fetchMock.mockResolvedValue(res(special));
    setOffice(makeDraft({ type: "text" }));
    render(<App />);
    await click("Sugerir resposta");
    await click(strings.composeInsert);
    const [data, opts] = item.body.prependAsync.mock.calls[0];
    expect(data).toBe(`${special}\n\n`);
    expect(opts).toEqual({ coercionType: "text" });
  });

  it("never uses another write or send API, and nothing happens without a click", async () => {
    setOffice(makeDraft());
    render(<App />);
    await click("Sugerir resposta");
    await click(strings.composeInsert);
    expect(item.body.setAsync).not.toHaveBeenCalled();
    expect(item.body.setSelectedDataAsync).not.toHaveBeenCalled();
    expect(item.displayReplyForm).not.toHaveBeenCalled();
    expect(item.displayReplyFormAsync).not.toHaveBeenCalled();
  });

  it.each([
    ["prependAsync fails", { prepend: { status: "failed" } as R }],
    ["prependAsync throws", { prepend: "throw" as const }],
  ])("%s: shows the alert, keeps the text, no fallback", async (_n, o) => {
    setOffice(makeDraft(o));
    render(<App />);
    await click("Sugerir resposta");
    await click(strings.composeInsert);
    const alert = screen.getByRole("alert");
    expect(alert.textContent).toBe(strings.composeInsertFailed);
    expect(screen.getByText("Bom dia Ana")).toBeTruthy();
    expect(item.body.prependAsync).toHaveBeenCalledTimes(1);
    expect(item.body.setAsync).not.toHaveBeenCalled();
    expect(btn(strings.composeInsert).disabled).toBe(false); // can try again
  });

  it("getTypeAsync failing shows the alert and writes nothing", async () => {
    setOffice(makeDraft());
    item.body.getTypeAsync.mockImplementation((cb: (r: R) => void) => cb({ status: "failed" }));
    render(<App />);
    await click("Sugerir resposta");
    await click(strings.composeInsert);
    expect(screen.getByRole("alert").textContent).toBe(strings.composeInsertFailed);
    expect(item.body.prependAsync).not.toHaveBeenCalled();
  });
});

describe("source guard (F-010/AC-13)", () => {
  const dir = join("src", "taskpane");
  const files = readdirSync(dir).filter((f) => /\.(ts|tsx)$/.test(f) && !/\.test\.|-DIM27/.test(f));
  const text = (f: string) => readFileSync(join(dir, f), "utf8");
  it("prependAsync is called only in composeInsert.ts; no other write or send API", () => {
    const forbidden =
      /setAsync|setSelectedDataAsync|setSignatureAsync|appendOnSendAsync|prependOnSendAsync|item\.saveAsync|addFileAttachmentAsync|makeEwsRequestAsync|displayReplyAllForm|ReadWriteMailbox|Mail\.Send|dangerouslySetInnerHTML|innerHTML|console\./;
    for (const f of files) {
      expect(text(f), f).not.toMatch(forbidden);
      if (f !== "composeInsert.ts") expect(text(f), f).not.toContain("prependAsync");
    }
    expect(text("composeInsert.ts").match(/\.prependAsync\(/g)).toHaveLength(1);
  });
  it("compose never reads recipients, sender or ids", () => {
    for (const f of ["composeItem.ts", "composeInsert.ts", "ComposeReply.tsx"])
      expect(text(f), f).not.toMatch(/\.(to|cc|bcc|from|sender|recipients|itemId)\b/);
  });
});
