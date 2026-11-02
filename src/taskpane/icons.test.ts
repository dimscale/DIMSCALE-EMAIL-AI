/* global process */
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

describe("add-in icons (F-012/AC-24)", () => {
  it.each([16, 32, 64, 80, 128])("icon-%i.png has the exact size and is 8-bit RGBA", (n) => {
    const b = readFileSync(join(process.cwd(), "assets", `icon-${n}.png`));
    expect(b.readUInt32BE(16)).toBe(n);
    expect(b.readUInt32BE(20)).toBe(n);
    expect([b[24], b[25]]).toEqual([8, 6]);
  });
});
