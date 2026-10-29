/* global process, Buffer */
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";

const SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

describe("add-in icons (F-012/AC-24)", () => {
  it.each([16, 32, 64, 80, 128])("icon-%i.png has the exact size and is 8-bit RGBA", (n) => {
    const b = readFileSync(join(process.cwd(), "assets", `icon-${n}.png`));
    expect(b.readUInt32BE(16)).toBe(n);
    expect(b.readUInt32BE(20)).toBe(n);
    expect([b[24], b[25]]).toEqual([8, 6]);
  });
});

describe("Resumo ribbon icons (F-012/AC-51)", () => {
  it.each([16, 32, 80])("resumo-%i.png is a PNG of the exact size, 8-bit RGBA", (n) => {
    const b = readFileSync(join(process.cwd(), "assets", `resumo-${n}.png`));
    expect([...b.subarray(0, 8)]).toEqual(SIGNATURE);
    expect(b.readUInt32BE(16)).toBe(n);
    expect(b.readUInt32BE(20)).toBe(n);
    expect([b[24], b[25]]).toEqual([8, 6]);
  });

  it("the master SVG has no text and the PNGs have transparent and opaque pixels", () => {
    const svg = readFileSync(join(process.cwd(), "assets", "icon-resumo-master.svg"), "utf8");
    expect(svg).not.toMatch(/<text|<rect/);
    for (const n of [16, 32, 80]) {
      const alphas = alphaChannel(
        readFileSync(join(process.cwd(), "assets", `resumo-${n}.png`)),
        n
      );
      expect(alphas).toContain(0);
      expect(Math.max(...alphas)).toBeGreaterThan(200);
      expect(alphas[0]).toBe(0); // corner is transparent
    }
  });
});

// Minimal decoder for 8-bit RGBA, non-interlaced PNGs: returns the alpha of every pixel.
function alphaChannel(b: Buffer, n: number): number[] {
  const idat: Buffer[] = [];
  for (let o = 8; o < b.length;) {
    const len = b.readUInt32BE(o);
    if (b.toString("latin1", o + 4, o + 8) === "IDAT") idat.push(b.subarray(o + 8, o + 8 + len));
    o += 12 + len;
  }
  const raw = inflateSync(Buffer.concat(idat));
  const stride = n * 4;
  const cur = Buffer.alloc(stride);
  let prev = Buffer.alloc(stride);
  const out: number[] = [];
  for (let y = 0; y < n; y++) {
    const f = raw[y * (stride + 1)];
    const line = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    for (let x = 0; x < stride; x++) {
      const a = x >= 4 ? cur[x - 4] : 0;
      const up = prev[x];
      const c = x >= 4 ? prev[x - 4] : 0;
      const p = a + up - c;
      const pa = Math.abs(p - a),
        pb = Math.abs(p - up),
        pc = Math.abs(p - c);
      const paeth = pa <= pb && pa <= pc ? a : pb <= pc ? up : c;
      cur[x] = (line[x] + [0, a, up, (a + up) >> 1, paeth][f]) & 255;
    }
    for (let x = 3; x < stride; x += 4) out.push(cur[x]);
    prev = Buffer.from(cur);
  }
  return out;
}
