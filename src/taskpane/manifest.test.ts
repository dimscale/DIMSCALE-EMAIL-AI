/* global process */

import { execSync } from "node:child_process";
import { mkdtempSync, readFileSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { describe, expect, it } from "vitest";

// Added in QA pass for F-001 (AC-2): static guard on manifest.xml. A text check, not a schema check.
// Relies on the repo root as cwd, which `npm test` sets.
const manifest = readFileSync(resolve(process.cwd(), "manifest.xml"), "utf8");

describe("manifest.xml (AC-2)", () => {
  it("requests only the ReadItem permission", () => {
    expect([...manifest.matchAll(/<Permissions>([^<]*)<\/Permissions>/g)].map((m) => m[1])).toEqual(
      ["ReadItem"]
    );
    expect(manifest).not.toMatch(/ReadWriteItem|ReadWriteMailbox|Mail\.Send/);
  });

  it("activates only on the message read form", () => {
    expect([...manifest.matchAll(/<Rule xsi:type="ItemIs"[^>]*\/>/g)].map((m) => m[0])).toEqual([
      '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />',
    ]);
    expect(manifest).not.toMatch(
      /MessageComposeCommandSurface|AppointmentOrganizer|AppointmentAttendee/
    );
  });

  it("declares no Graph/SSO web application info and no send-related events", () => {
    expect(manifest).not.toMatch(/WebApplicationInfo|ItemSend|ItemInsert/);
  });

  it("supports pinning and has the DIMSCALE MAIL group (F-012)", () => {
    expect(manifest).toContain('id="GroupLabel" DefaultValue="DIMSCALE MAIL"');
    expect(manifest).not.toContain("CustomTab");
  });
});

describe.each(["manifest.xml", "manifest.customtab-test.xml"])(
  "%s ribbon buttons (F-012/AC-48)",
  (file) => {
    const m = readFileSync(resolve(process.cwd(), file), "utf8");
    const buttons = [
      ...m.matchAll(/<Control xsi:type="Button" id="([^"]+)">([\s\S]*?)<\/Control>/g),
    ];

    it("has exactly one button, DIMSCALE Mail, a pinnable ShowTaskpane", () => {
      expect(buttons.map((b) => b[1])).toEqual(["msgReadOpenPaneButton"]);
      expect(buttons[0][2]).toContain('<Action xsi:type="ShowTaskpane">');
      expect(buttons[0][2]).toContain("<SupportsPinning>true</SupportsPinning>");
      expect(m).toContain('id="IaButton.Label" DefaultValue="DIMSCALE Mail"');
      expect(m).not.toContain("ResumoButton");
    });

    it("points at taskpane.html only", () => {
      expect(buttons[0][2]).toContain('<SourceLocation resid="Taskpane.Url" />');
      expect(m).toContain(
        '<bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />'
      );
      expect(m).not.toContain("resumo.html");
    });

    it("every resid is defined, resource ids are at most 32 chars, only ReadItem", () => {
      const defined = new Set(
        [...m.matchAll(/<bt:(?:Image|Url|String) id="([^"]+)"/g)].map((x) => x[1])
      );
      for (const r of m.matchAll(/resid="([^"]+)"/g)) expect(defined.has(r[1]), r[1]).toBe(true);
      for (const id of defined) expect(id.length).toBeLessThanOrEqual(32);
      expect([...m.matchAll(/<Permissions>([^<]*)<\/Permissions>/g)].map((x) => x[1])).toEqual([
        "ReadItem",
      ]);
      expect(m).not.toContain("<AppDomains");
    });

    it("uses the green D icons", () => {
      for (const n of [16, 32, 80]) {
        expect(m).toContain(
          `id="Icon.Ia.${n}x${n}" DefaultValue="https://localhost:3000/assets/icon-${n}.png"`
        );
      }
    });
  }
);

it("versions: manifest 1.0.2.0, test manifest 1.0.2.1", () => {
  expect(manifest).toContain("<Version>1.0.2.0</Version>");
  expect(readFileSync(resolve(process.cwd(), "manifest.customtab-test.xml"), "utf8")).toContain(
    "<Version>1.0.2.1</Version>"
  );
});

describe("production build (F-012/AC-48)", () => {
  it("emits taskpane.html and resumo.html with the same hashed assets and rewritten URLs", () => {
    const out = mkdtempSync(join(tmpdir(), "dmai-build-"));
    execSync(`npx webpack --mode production --output-path "${out}"`, {
      cwd: process.cwd(),
      stdio: "pipe",
    });
    expect(readdirSync(out)).toEqual(expect.arrayContaining(["taskpane.html", "resumo.html"]));
    const assets = (f: string) =>
      [...readFileSync(join(out, f), "utf8").matchAll(/(?:src|href)="([^"]+\.js)"/g)]
        .map((x) => x[1])
        .filter((x) => !x.startsWith("https://")) // skip the external office.js CDN script
        .sort();
    const a = assets("taskpane.html");
    expect(a.length).toBeGreaterThan(0);
    expect(a.every((x) => /\.[0-9a-f]{20}\.js$/.test(x))).toBe(true);
    expect(assets("resumo.html")).toEqual(a);
    const prod = readFileSync(join(out, "manifest.xml"), "utf8");
    expect(prod).toContain("/taskpane.html");
    expect(prod).not.toContain("resumo.html"); // the Resumo button was removed (user, 2026-10-29); the page stays built but is not linked
    expect(prod).not.toContain("localhost:3000");
    expect(readdirSync(out)).not.toContain("manifest.customtab-test.xml"); // AC-49: test manifest is not published
  }, 240000);
});
