/* global __dirname */
import { readdirSync, readFileSync } from "fs";
import { join } from "path";
import { describe, expect, it } from "vitest";

const dir = __dirname;
const files = readdirSync(dir).filter((f) => /\.tsx?$/.test(f) && !f.includes(".test."));
const src = files.map((f) => readFileSync(join(dir, f), "utf8")).join("\n");

describe("F-013 client hygiene", () => {
  it("requests no Graph scope other than Mail.ReadBasic", () => {
    const found = src.match(/graph\.microsoft\.com\/[A-Za-z.]+/g) ?? [];
    const scopes = found.filter((s) => /^graph\.microsoft\.com\/[A-Z]/.test(s));
    expect([...new Set(scopes)]).toEqual(["graph.microsoft.com/Mail.ReadBasic"]);
    expect(src).not.toMatch(/Mail\.Send|Mail\.ReadWrite|Contacts\.Read/);
  });
  it("never logs from the Responder modules", () => {
    for (const f of ["graph.ts", "settingsClient.ts", "ResponderPanel.tsx"]) {
      expect(readFileSync(join(dir, f), "utf8")).not.toMatch(/console\./);
    }
  });
});
