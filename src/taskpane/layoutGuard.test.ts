import { readFileSync } from "fs";
import { join } from "path";
import { describe, expect, it } from "vitest";

// F-012 layout guard: the pane must reflow down to 320 px. Static scan of the makeStyles sources.
const files = [
  "App.tsx",
  "MessagePanel.tsx",
  "AiTools.tsx",
  "ResponderPanel.tsx",
  "BackendStatus.tsx",
];
const src = (f: string) => readFileSync(join("src", "taskpane", f), "utf8");

describe("layout guard (320 px)", () => {
  it("has no white-space nowrap/pre outside the visually-hidden helper", () => {
    for (const f of files) {
      const t = src(f).replace(/hidden: \{[^}]*\}/, "");
      expect(t, f).not.toMatch(/whiteSpace:\s*"(nowrap|pre)"/);
    }
  });
  it("has no fixed width/minWidth above 280 px", () => {
    for (const f of files) {
      for (const m of src(f).matchAll(/\b(?:width|minWidth):\s*"(\d+)px"/g)) {
        expect(Number(m[1]), `${f} ${m[0]}`).toBeLessThanOrEqual(280);
      }
    }
  });
});
