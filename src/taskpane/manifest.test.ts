/* global process */

import { execSync } from "node:child_process";
import { mkdtempSync, readFileSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { describe, expect, it } from "vitest";

// Added in QA pass for F-001 (AC-2): static guard on manifest.xml. A text check, not a schema check.
// Relies on the repo root as cwd, which `npm test` sets.
const manifest = readFileSync(resolve(process.cwd(), "manifest.xml"), "utf8");

describe("manifest.xml (AC-2, F-010/AC-1, AC-2)", () => {
  it("requests only the ReadWriteItem permission, no send or mailbox-wide permission (F-010/AC-1)", () => {
    expect([...manifest.matchAll(/<Permissions>([^<]*)<\/Permissions>/g)].map((m) => m[1])).toEqual(
      ["ReadWriteItem"]
    );
    expect(manifest).not.toMatch(
      /ReadWriteMailbox|Mail\.Send|ExtendedPermissions|AppendOnSend|LaunchEvent|ItemSend|WebApplicationInfo/
    );
  });

  it("keeps the read rule and adds only the compose ribbon surface (F-010/AC-2)", () => {
    expect([...manifest.matchAll(/<Rule xsi:type="ItemIs"[^>]*\/>/g)].map((m) => m[0])).toEqual([
      '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />',
    ]);
    expect([...manifest.matchAll(/<ExtensionPoint xsi:type="([^"]+)">/g)].map((m) => m[1])).toEqual(
      ["MessageReadCommandSurface", "MessageComposeCommandSurface"]
    );
    expect(manifest).not.toMatch(
      /AppointmentOrganizer|AppointmentAttendee|ExecuteFunction|FunctionFile/
    );
  });

  it("declares no Graph/SSO web application info and no send-related events", () => {
    expect(manifest).not.toMatch(/WebApplicationInfo|ItemSend|ItemInsert/);
  });

  it("supports pinning and has the DIMSCALE group (F-012)", () => {
    expect(manifest).toContain('id="GroupLabel" DefaultValue="DIMSCALE"');
    expect(manifest).not.toContain("CustomTab");
  });
});

// The test-only manifest stays on the read surface and ReadItem (F-010/AC-16: it stays dropped, untouched).
describe.each([
  ["manifest.xml", ["msgReadOpenPaneButton", "msgComposeOpenPaneButton"], "ReadWriteItem"],
  ["manifest.customtab-test.xml", ["msgReadOpenPaneButton"], "ReadItem"],
])("%s ribbon buttons (F-012/AC-48, F-010/AC-2)", (file, ids, permission) => {
  const m = readFileSync(resolve(process.cwd(), file), "utf8");
  const buttons = [...m.matchAll(/<Control xsi:type="Button" id="([^"]+)">([\s\S]*?)<\/Control>/g)];

  it("has one DIMSCALE Mail button per surface, each a pinnable ShowTaskpane", () => {
    expect(buttons.map((b) => b[1])).toEqual(ids);
    for (const b of buttons) {
      expect(b[2]).toContain('<Action xsi:type="ShowTaskpane">');
      expect(b[2]).toContain("<SupportsPinning>true</SupportsPinning>");
      expect(b[2]).toContain('<Label resid="IaButton.Label" />');
    }
    expect(m).toContain('id="IaButton.Label" DefaultValue="DIMSCALE Mail"');
    expect(m).not.toContain("ResumoButton");
  });

  it("points at taskpane.html only", () => {
    for (const b of buttons) expect(b[2]).toContain('<SourceLocation resid="Taskpane.Url" />');
    expect(m).toContain(
      '<bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />'
    );
    expect(m).not.toContain("resumo.html");
  });

  it("every resid is defined, resource ids are at most 32 chars, one permission", () => {
    const defined = new Set(
      [...m.matchAll(/<bt:(?:Image|Url|String) id="([^"]+)"/g)].map((x) => x[1])
    );
    for (const r of m.matchAll(/resid="([^"]+)"/g)) expect(defined.has(r[1]), r[1]).toBe(true);
    for (const id of defined) expect(id.length).toBeLessThanOrEqual(32);
    expect([...m.matchAll(/<Permissions>([^<]*)<\/Permissions>/g)].map((x) => x[1])).toEqual([
      permission,
    ]);
    expect(m).not.toContain("<AppDomains");
  });

  it("uses the brand icons (DIMSCALE Λ since 2026-11-17)", () => {
    for (const n of [16, 32, 80]) {
      expect(m).toContain(
        `id="Icon.Ia.${n}x${n}" DefaultValue="https://localhost:3000/assets/icon-${n}.png"`
      );
    }
  });
});

it("versions and names: manifest 1.1.0.0, test manifest 1.0.4.1, DisplayName DIMSCALE Mail", () => {
  const test = readFileSync(resolve(process.cwd(), "manifest.customtab-test.xml"), "utf8");
  expect(manifest).toContain("<Version>1.1.0.0</Version>");
  expect(manifest).toContain('<DisplayName DefaultValue="DIMSCALE Mail"/>');
  expect(test).toContain("<Version>1.0.4.1</Version>");
  expect(test).toContain('<DisplayName DefaultValue="DIMSCALE Mail (teste separador)"/>');
  for (const m of [manifest, test]) expect(m).not.toMatch(/Assistente IA|DIMSCALE Mail AI/);
});

describe("production build (F-012/AC-48)", () => {
  it("emits taskpane.html with hashed assets and rewritten URLs, and no resumo.html", () => {
    const out = mkdtempSync(join(tmpdir(), "dmai-build-"));
    execSync(`npx webpack --mode production --output-path "${out}"`, {
      cwd: process.cwd(),
      stdio: "pipe",
      // the default build must not inherit the Cloudflare build environment (this suite also runs there)
      env: { ...process.env, CF_PAGES_BRANCH: "", PANE_BASE_URL: "", PANE_TEST: "" },
    });
    expect(readdirSync(out)).toContain("taskpane.html");
    expect(readdirSync(out)).not.toContain("resumo.html"); // the Resumo ribbon button was removed (user, 2026-10-29)
    const assets = (f: string) =>
      [...readFileSync(join(out, f), "utf8").matchAll(/(?:src|href)="([^"]+\.js)"/g)]
        .map((x) => x[1])
        .filter((x) => !x.startsWith("https://")) // skip the external office.js CDN script
        .sort();
    const a = assets("taskpane.html");
    expect(a.length).toBeGreaterThan(0);
    expect(a.every((x) => /\.[0-9a-f]{20}\.js$/.test(x))).toBe(true);
    const prod = readFileSync(join(out, "manifest.xml"), "utf8");
    // ADR-0007: default host stays GitHub Pages; the page is referenced without .html (Cloudflare redirects *.html)
    expect(prod).toContain('DefaultValue="https://dimscale.github.io/DIMSCALE-EMAIL-AI/taskpane"');
    expect(prod).not.toContain("taskpane.html");
    expect(prod).toContain("<Id>05c2e1c9-3e1d-406e-9a91-e9ac64854143</Id>");
    expect(prod).not.toContain("resumo.html");
    expect(prod).not.toContain("localhost:3000");
    expect(prod).toContain("<Version>1.1.0.0</Version>"); // F-010/AC-3
    expect(prod).toContain("<Permissions>ReadWriteItem</Permissions>");
    expect(prod).toContain('xsi:type="MessageReadCommandSurface"');
    expect(prod).toContain('xsi:type="MessageComposeCommandSurface"');
    expect(readdirSync(out)).not.toContain("manifest.customtab-test.xml"); // AC-49: test manifest is not published
  }, 240000);

  const build = (env: Record<string, string>) => {
    const out = mkdtempSync(join(tmpdir(), "dmai-build-"));
    execSync(`npx webpack --mode production --output-path "${out}"`, {
      cwd: process.cwd(),
      stdio: "pipe",
      env: { ...process.env, ...env },
    });
    return readFileSync(join(out, "manifest.xml"), "utf8");
  };

  it("ADR-0007: PANE_BASE_URL outside the closed list blocks the build", () => {
    expect(() => build({ PANE_BASE_URL: "https://evil.example/" })).toThrow();
  }, 240000);

  it("ADR-0007: a Cloudflare build of the dev branch is the test add-in without manual variables", () => {
    const m = build({ CF_PAGES_BRANCH: "dev", PANE_BASE_URL: "", PANE_TEST: "" });
    expect(m).toContain("<Id>926a7aa2-46e8-4187-9111-c9873e1fdec5</Id>");
    expect(m).toContain('DefaultValue="https://dev.dimscale-mail-ai.pages.dev/taskpane"');
    const main = build({ CF_PAGES_BRANCH: "main", PANE_BASE_URL: "", PANE_TEST: "" });
    expect(main).toContain("<Id>05c2e1c9-3e1d-406e-9a91-e9ac64854143</Id>");
    expect(main).not.toContain("(teste)");
    expect(main).toContain('DefaultValue="https://mailai.dimscale.com/taskpane"'); // Cloudflare production
  }, 240000);

  it("ADR-0007: PANE_TEST preview build gets its own Id, name and host", () => {
    const m = build({ PANE_BASE_URL: "https://dev.dimscale-mail-ai.pages.dev/", PANE_TEST: "1" });
    expect(m).toContain("<Id>926a7aa2-46e8-4187-9111-c9873e1fdec5</Id>");
    expect(m).not.toContain("05c2e1c9-3e1d-406e-9a91-e9ac64854143");
    expect(m).toContain('<DisplayName DefaultValue="DIMSCALE Mail (teste)"/>');
    expect(m).toContain('id="IaButton.Label" DefaultValue="DIMSCALE Mail (teste)"'); // the ribbon button is told apart too
    // the compose button uses the same IaButton.Label resid, so it carries "(teste)" as well (F-010)
    expect(m).toContain('xsi:type="MessageComposeCommandSurface"');
    expect(m.match(/<Label resid="IaButton.Label" \/>/g)).toHaveLength(2);
    expect(m).toContain('DefaultValue="https://dev.dimscale-mail-ai.pages.dev/taskpane"');
    expect(m).not.toContain("dimscale.github.io");
  }, 240000);
});
