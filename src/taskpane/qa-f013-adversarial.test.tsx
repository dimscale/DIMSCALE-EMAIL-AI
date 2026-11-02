// QA (adversarial) for F-013 amendment AC-33, AC-43, AC-49 (pane side). Mocked Graph, settings and auth; no live Outlook.
import * as React from "react";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";

vi.mock("./auth", () => ({
  getTokenSilent: vi.fn(),
  isNaaSupported: vi.fn(() => true),
}));

// Keep the real pure helpers (filterRecentSenders, isNoReplySender); mock only the Graph token and fetch calls.
vi.mock("./graph", async (importOriginal) => ({
  ...(await importOriginal<typeof import("./graph")>()),
  getGraphTokenSilent: vi.fn(),
  requestGraphConsent: vi.fn(),
  fetchRecentSenders: vi.fn(),
}));

vi.mock("./settingsClient", () => ({
  fetchSettings: vi.fn(),
  saveSettings: vi.fn(),
  deleteSettings: vi.fn(),
}));

import { getTokenSilent } from "./auth";
import { filterRecentSenders, fetchRecentSenders, getGraphTokenSilent, isNoReplySender } from "./graph";
import { fetchSettings } from "./settingsClient";
import ResponderPanel from "./ResponderPanel";

const NOREPLY_MSG = "Não pode adicionar um endereço de não-resposta (por exemplo noreply).";
const read = (p: string) => readFileSync(join(process.cwd(), p), "utf8");

afterEach(cleanup);

describe("F-013 candidates (AC-33) with the real helpers", () => {
  it("a display name that looks like a no-reply address does not exclude a real sender", () => {
    expect(filterRecentSenders([{ from: { emailAddress: { address: "ana@x.pt", name: "noreply@evil.com" } } }])).toEqual([
      { address: "ana@x.pt", name: "noreply@evil.com" },
    ]);
  });

  it("a no-reply address is excluded whatever the display name, in any case and with spaces", () => {
    expect(filterRecentSenders([{ from: { emailAddress: { address: " NoReply@X.pt ", name: "Ana Silva" } } }])).toEqual([]);
    expect(filterRecentSenders([{ from: { emailAddress: { address: "Postmaster@x.pt", name: "" } } }])).toEqual([]);
  });

  it("duplicates after trim and case collapse to one candidate; own address excluded", () => {
    const r = filterRecentSenders(
      [
        { from: { emailAddress: { address: "ana@x.pt" } } },
        { from: { emailAddress: { address: " ANA@x.pt " } } },
        { from: { emailAddress: { address: "owner@dimscale.com" } } },
      ],
      "OWNER@dimscale.com"
    );
    expect(r.map((c) => c.address)).toEqual(["ana@x.pt"]);
  });

  it("plus-addressing, fullwidth and zero-width variants are excluded from candidates; noreply.team and a Cyrillic look-alike are not", () => {
    for (const a of ["noreply+x@x.pt", "ｎｏｒｅｐｌｙ@x.pt", "noreply​@x.pt"]) expect(isNoReplySender(a)).toBe(true);
    expect(isNoReplySender("noreply.team@x.pt")).toBe(false);
    expect(isNoReplySender("nоreply@x.pt")).toBe(false); // known limit (Cyrillic o)
  });
});

describe("F-013 no-reply rule parity (AC-33, AC-49): pane and function use the same rule", () => {
  it("src/taskpane/noReply.ts and supabase/functions/_shared/filters.ts hold the same list and the same normalisation", () => {
    const list = /const NOREPLY = (\/\^\(.*?\)\$\/);/;
    const strip = /\.replace\((\/\[\\p\{Default_Ignorable_Code_Point\}.*?\/gu), ""\)/;
    const pane = read("src/taskpane/noReply.ts");
    const fn = read("supabase/functions/_shared/filters.ts");
    expect(pane.match(list)?.[1]).toBeTruthy();
    expect(pane.match(list)?.[1]).toBe(fn.match(list)?.[1]);
    expect(pane.match(strip)?.[1]).toBeTruthy();
    expect(pane.match(strip)?.[1]).toBe(fn.match(strip)?.[1]);
    for (const src of [pane, fn]) expect(src).toContain('.normalize("NFKC")');
    for (const src of [pane, fn]) expect(src).toContain('.split("+")[0]');
  });
});

describe("F-013 manual entry (AC-49, pane side)", () => {
  beforeEach(() => {
    (globalThis as any).Office = {
      context: { mailbox: { userProfile: { emailAddress: "thiago.santos@dimscale.com" } } },
    };
    (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "mock_backend_token" });
    (fetchSettings as any).mockResolvedValue({
      kind: "ok",
      data: { enabled: true, status: "active", contacts: ["ana@cliente.pt"] },
    });
    (getGraphTokenSilent as any).mockResolvedValue({ kind: "token", token: "mock_graph_token" });
    (fetchRecentSenders as any).mockResolvedValue([{ address: "rui@empresa.com", name: "Rui Costa" }]);
  });

  it("refuses padded and mixed-case no-reply addresses with the text message; a plus-addressed one is refused too", async () => {
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    const input = screen.getByRole("textbox", { name: "Adicionar endereço" });
    const add = screen.getByRole("button", { name: "Adicionar" });

    for (const v of ["  NoReply@X.pt  ", "\tpostmaster@x.pt", "NOREPLY@x.pt", "noreply+x@x.pt"]) {
      fireEvent.change(input, { target: { value: v } });
      fireEvent.click(add);
      expect(screen.getByText(NOREPLY_MSG)).toBeTruthy();
      expect(screen.getByText("Contactos a responder (1)")).toBeTruthy();
    }
  });

  it("a display-name form is refused as an invalid address (no-reply check is not reached, nothing added)", async () => {
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    const input = screen.getByRole("textbox", { name: "Adicionar endereço" });
    fireEvent.change(input, { target: { value: "NoReply <noreply@x.pt>" } });
    fireEvent.click(screen.getByRole("button", { name: "Adicionar" }));
    expect(screen.getByText("Endereço de email inválido.")).toBeTruthy();
    expect(screen.getByText("Contactos a responder (1)")).toBeTruthy();
  });
});
