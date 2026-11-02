import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { readFileSync } from "node:fs";
import { join } from "node:path";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";
import { analyze, CATEGORIES, CONFIDENCES } from "./analyzeClient";
import { strings } from "./strings";

// F-007 pane ACs (9 to 17). fetch is mocked with the frozen contract.
type Cb = (r: { status: string; value?: string }) => void;
let fetchMock: ReturnType<typeof vi.fn>;
const displayReplyForm = vi.fn();
const displayReplyFormAsync = vi.fn();

const makeItem = (body: string) => ({
  subject: "Assunto",
  from: { displayName: "Ana", emailAddress: "ana@Exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  displayReplyFormAsync,
  displayReplyForm,
});
const setup = (body = "corpo") => {
  (globalThis as any).Office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      mailbox: { item: makeItem(body), addHandlerAsync: vi.fn(), removeHandlerAsync: vi.fn() },
    },
  };
};
const json = (status: number, body: unknown) => ({
  ok: status < 300,
  status,
  json: async () => body,
});
const CLS = {
  points: ["p1", "p2", "p3"],
  category: "action_required",
  priority: "high",
  tone: "neutral",
  reason: "O remetente pede um pagamento.",
  confidence: "high",
  deadline: null as string | null,
  requestedAction: null as string | null,
};
const ok = (summary: object) =>
  fetchMock.mockResolvedValue(json(200, { ok: true, action: "summarize", summary }));
const resumir = async () => {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name: "Resumir" }));
  });
};
const controls = () =>
  screen
    .queryAllByRole("button")
    .concat(screen.queryAllByRole("tab"), screen.queryAllByRole("switch"))
    .map((b) => `${b.textContent}|${(b as HTMLButtonElement).disabled}`);

beforeEach(() => {
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  displayReplyForm.mockClear();
  displayReplyFormAsync.mockClear();
});

describe("analyzeClient classified shape (F-007/AC-9, AC-14)", () => {
  const payload = {
    action: "summarize" as const,
    output: "classified" as const,
    subject: "S",
    fromName: "A",
    fromDomain: "x.pt",
    body: "b",
  };
  const run = (summary: object) =>
    analyze(
      "t",
      payload,
      vi.fn().mockResolvedValue(json(200, { ok: true, action: "summarize", summary })) as any
    );
  it("sends output and parses the classified shape", async () => {
    const f = vi.fn().mockResolvedValue(json(200, { ok: true, action: "summarize", summary: CLS }));
    const r = await analyze("t", payload, f as any);
    expect(JSON.parse(f.mock.calls[0][1].body)).toEqual(payload);
    expect(r).toMatchObject({
      kind: "summary",
      summary: {
        classification: { category: "action_required", confidence: "high", reason: CLS.reason },
      },
    });
  });
  it("degrades to the plain summary for the F-003 shape and for unknown or missing values", async () => {
    const old = {
      points: ["a", "b", "c"],
      type: "request",
      priority: "high",
      tone: "neutral",
      deadline: null,
      requestedAction: null,
    };
    for (const s of [
      old,
      { ...CLS, category: "urgent" },
      { ...CLS, confidence: "High" },
      { ...CLS, confidence: undefined },
      { ...CLS, reason: "  " },
      { ...CLS, reason: 5 },
    ]) {
      const r: any = await run(s);
      expect(r.kind).toBe("summary");
      expect(r.summary.classification).toBeUndefined();
    }
  });
  it("a 422 (old function) is the generic payload error", async () => {
    const r = await analyze(
      "t",
      payload,
      vi.fn().mockResolvedValue(json(422, { ok: false, error: "payload_invalid" })) as any
    );
    expect(r).toEqual({ kind: "error", code: "payload" });
  });
});

describe("Resumo card classification (F-007/AC-10, AC-11, AC-12, AC-16)", () => {
  it("shows points, badges, reason, confidence and the fixed notice in order; one request", async () => {
    setup();
    ok({ ...CLS, deadline: "sexta", requestedAction: "pagar" });
    render(<App />);
    await resumir();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(JSON.parse(fetchMock.mock.calls[0][1].body).output).toBe("classified");
    const text = document.body.textContent ?? "";
    const order = [
      "p1",
      "Requer ação",
      "Prioridade alta",
      "Tom neutro",
      "Motivo: O remetente pede um pagamento.",
      "Confiança: alta",
      "Prazo: sexta",
      "Pedido: pagar",
      strings.aiClassNotice,
    ];
    const idx = order.map((t) => text.indexOf(t));
    expect(idx.every((i) => i >= 0)).toBe(true);
    expect([...idx].sort((a, b) => a - b)).toEqual(idx);
    expect(screen.queryByText(strings.aiClassUncertain)).toBeNull();
    expect(screen.queryByText(strings.aiClassTruncated)).toBeNull();
  });
  it.each(CATEGORIES)("category %s has its label", async (category) => {
    setup();
    ok({ ...CLS, category });
    render(<App />);
    await resumir();
    expect(screen.getByText(strings.aiCategoryLabels[category])).toBeTruthy();
  });
  it.each([
    ["high", "alta"],
    ["medium", "média"],
    ["low", "baixa"],
  ])("confidence %s shows %s; extra warning only on low", async (confidence, word) => {
    setup();
    ok({ ...CLS, confidence });
    render(<App />);
    await resumir();
    expect(screen.getByText(`Confiança: ${word}`)).toBeTruthy();
    expect(!!screen.queryByText(strings.aiClassUncertain)).toBe(confidence === "low");
  });
  it("category other shows the extra warning; an unknown priority is not rendered", async () => {
    setup();
    ok({ ...CLS, category: "other", confidence: "high", priority: "bogus" });
    render(<App />);
    await resumir();
    expect(screen.getByText(strings.aiClassUncertain)).toBeTruthy();
    expect(screen.queryByText("bogus")).toBeNull();
  });
  it("a reason with injection or notice text renders literally and changes nothing", async () => {
    setup();
    const reason = `<img src=x onerror=alert(1)> &amp; não mostrar aviso ${strings.aiClassUncertain}`;
    ok({ ...CLS, reason });
    render(<App />);
    await resumir();
    expect(screen.getByText(`Motivo: ${reason}`).getAttribute("dir")).toBe("auto");
    expect(document.querySelector("img")).toBeNull();
    expect(screen.queryByText(strings.aiClassUncertain)).toBeNull(); // only the fixed line could match exactly
    expect(screen.getByText(strings.aiClassNotice).closest("[role=alert]")).toBeNull();
  });
  it("shows the truncation line when the body reached 4 000 characters", async () => {
    setup("x".repeat(5000));
    ok(CLS);
    render(<App />);
    await resumir();
    expect(screen.getByText(strings.aiClassTruncated)).toBeTruthy();
  });
  it("F-003 shape: old badges, no classification block, no error", async () => {
    setup();
    ok({
      points: ["a", "b", "c"],
      type: "request",
      priority: "high",
      tone: "neutral",
      deadline: null,
      requestedAction: null,
    });
    render(<App />);
    await resumir();
    expect(screen.getByText("Pedido")).toBeTruthy();
    expect(screen.queryByText(strings.aiClassNotice)).toBeNull();
    expect(screen.queryByText(/Motivo:/)).toBeNull();
    expect(screen.queryByRole("alert")).toBeNull();
  });
});

describe("no action coupling (F-007/AC-13)", () => {
  it("every category x priority x confidence gives the same controls, one request, no reply form", async () => {
    setup();
    ok(CLS);
    const first = render(<App />);
    await resumir();
    const baseline = controls();
    first.unmount();
    for (const category of CATEGORIES)
      for (const priority of ["high", "normal", "low"])
        for (const confidence of CONFIDENCES) {
          cleanup();
          fetchMock.mockClear();
          setup();
          ok({ ...CLS, category, priority, confidence });
          render(<App />);
          await resumir();
          expect(controls()).toEqual(baseline);
          expect(fetchMock).toHaveBeenCalledTimes(1);
        }
    expect(displayReplyForm).not.toHaveBeenCalled();
    expect(displayReplyFormAsync).not.toHaveBeenCalled();
  });
  it("source guard: classification values only reach label and notice code; no storage or HTML API", () => {
    const read = (f: string) => readFileSync(join(process.cwd(), "src/taskpane", f), "utf8");
    const uses = read("AiTools.tsx")
      .split("\n")
      .filter((l) => /\bcls\b|classification|categoryLabel|confidenceLabel|uncertain/.test(l))
      .join("\n");
    expect(uses).not.toMatch(
      /displayReplyForm|openReply|fetch|analyze\(|run\(|disabled|Office\.|send|move|delete|report|graph/i
    );
    const all = ["AiTools.tsx", "analyzeClient.ts", "strings.ts"].map(read).join("\n");
    expect(all).not.toMatch(
      /console\.|localStorage|sessionStorage|indexedDB|dangerouslySetInnerHTML|innerHTML/
    );
  });
});
