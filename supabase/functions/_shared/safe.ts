const enc = new TextEncoder();

export async function sha256Hex(s: string): Promise<string> {
  const d = new Uint8Array(await crypto.subtle.digest("SHA-256", enc.encode(s)));
  return Array.from(d, (b) => b.toString(16).padStart(2, "0")).join("");
}

/** Constant-time string compare: both sides are hashed first, so length does not leak. */
export async function ctEqual(a: string | undefined | null, b: string | undefined | null): Promise<boolean> {
  if (!a || !b) return false;
  const [x, y] = await Promise.all([crypto.subtle.digest("SHA-256", enc.encode(a)), crypto.subtle.digest("SHA-256", enc.encode(b))]);
  const xa = new Uint8Array(x), ya = new Uint8Array(y);
  let diff = 0;
  for (let i = 0; i < xa.length; i++) diff |= xa[i] ^ ya[i];
  return diff === 0;
}

export const normalizeAddress = (s: string) => s.trim().toLowerCase();
