import * as React from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
import App from "./App";

afterEach(cleanup);

describe("App layout (F-012)", () => {
  it("has one h1, tabs Analisar/Responder, the assurance line and the coming-soon cards", () => {
    (globalThis as any).Office = { CoercionType: { Text: "text" }, context: { mailbox: { item: undefined } } };
    render(<App />);
    expect(screen.getAllByRole("heading", { level: 1 })).toHaveLength(1);
    expect(screen.getByRole("img", { name: "DIMSCALE" })).toBeTruthy();
    expect(screen.getByRole("tab", { name: "Analisar" }).getAttribute("aria-selected")).toBe("true");
    expect(screen.getByText("Sem envio automático. O envio é sempre feito por si.")).toBeTruthy();
    expect(screen.getByText("IA (resumo e resposta sugerida) em breve.")).toBeTruthy();
    fireEvent.click(screen.getByRole("tab", { name: "Responder" }));
    expect(screen.getByText("Respostas automáticas para contactos escolhidos: em breve.")).toBeTruthy();
    expect(screen.getByRole("tabpanel", { name: "Responder" })).toBeTruthy();
    expect(screen.queryByRole("tabpanel", { name: "Analisar" })).toBeNull(); // hidden panel is out of the a11y tree
  });
});

describe("source guards (F-012/AC-3, AC-4)", () => {
  const dir = join(process.cwd(), "src/taskpane");
  const files = readdirSync(dir).filter((f) => /\.tsx?$/.test(f) && !/\.test\./.test(f));
  const src = files.map((f) => readFileSync(join(dir, f), "utf8")).join("\n");
  it("uses no HTML injection, storage or console", () => {
    expect(src).not.toMatch(/dangerouslySetInnerHTML|innerHTML|insertAdjacentHTML/);
    expect(src).not.toMatch(/localStorage|sessionStorage|document\.cookie|indexedDB|console\./);
    expect(src).not.toMatch(/isDarkTheme|themeId/);
  });
});

describe("backend notice placement (F-012)", () => {
  it("renders the BackendStatus notice in the host above the tab content and hides the footer line from AT", async () => {
    vi.resetModules();
    vi.doUnmock("./BackendStatus");
    vi.doMock("./config", () => ({ isConfigured: () => false }));
    vi.doMock("./auth", () => ({ isNaaSupported: () => true, getTokenSilent: vi.fn(), signIn: vi.fn() }));
    const { default: RealApp } = await import("./App");
    (globalThis as any).Office = { CoercionType: { Text: "text" }, context: { mailbox: { item: undefined } } };
    const { container } = render(<RealApp />);
    const bar = await screen.findByText(/ainda não está configurado/);
    expect(container.querySelector("footer")?.contains(bar)).toBe(false);
    expect(screen.queryByRole("status", { name: "" , hidden: false })).toBeNull();
    expect(container.querySelector('footer [aria-hidden="true"]')?.textContent).toContain("Backend não configurado");
  });
});
