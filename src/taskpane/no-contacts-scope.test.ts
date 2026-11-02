/* global process */
import { describe, expect, it } from "vitest";
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";

// F-013/AC-43: no contacts scope in src/ or the manifests. Needles are split so this file does not match itself.
const NEEDLES = [
  "Contacts" + ".Read",
  "/me/" + "contacts",
  "/me/" + "people",
  "Mailbox" + "Settings",
];

const walk = (dir: string): string[] =>
  readdirSync(dir).flatMap((n) => {
    const p = join(dir, n);
    return statSync(p).isDirectory() ? walk(p) : [p];
  });

describe("F-013/AC-43 no contacts scope", () => {
  it("src/ and manifest*.xml never mention contacts, people or mailbox-settings scopes", () => {
    const root = process.cwd();
    const files = [
      ...walk(join(root, "src")).filter((f) => /\.(tsx?|json|html|css)$/.test(f)),
      ...readdirSync(root)
        .filter((n) => /^manifest.*\.xml$/.test(n))
        .map((n) => join(root, n)),
    ];
    expect(files.length).toBeGreaterThan(5);
    const hits = files.filter((f) => {
      const t = readFileSync(f, "utf8");
      return NEEDLES.some((n) => t.includes(n));
    });
    expect(hits).toEqual([]);
  });
});
