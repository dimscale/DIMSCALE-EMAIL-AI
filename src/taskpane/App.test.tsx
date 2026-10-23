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
    expect(screen.getByText("Sem envio automático")).toBeTruthy();
    expect(screen.getByText("Resumo e resposta sugerida — em breve")).toBeTruthy();
    fireEvent.click(screen.getByRole("tab", { name: "Responder" }));
    expect(screen.getByText("Respostas automáticas para contactos selecionados — em breve")).toBeTruthy();
    expect(screen.queryByText("Sem envio automático")).toBeNull();
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
