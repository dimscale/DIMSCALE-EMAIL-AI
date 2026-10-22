/* global fetch, setTimeout, clearTimeout, AbortController */
import { FUNCTION_URL } from "./config";

export type WhoamiResult =
  | { kind: "ok"; name: string | null }
  | { kind: "rejected" }
  | { kind: "unreachable"; status?: number };

export async function callWhoami(token: string): Promise<WhoamiResult> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 10000);
  try {
    const res = await fetch(FUNCTION_URL, {
      method: "GET",
      headers: { Authorization: `Bearer ${token}` },
      signal: ctrl.signal,
    });
    if (res.status === 401) return { kind: "rejected" };
    if (res.status !== 200) return { kind: "unreachable", status: res.status };
    const body = await res.json();
    if (body?.ok !== true) return { kind: "unreachable", status: res.status };
    return { kind: "ok", name: typeof body.name === "string" && body.name ? body.name : null };
  } catch {
    return { kind: "unreachable" };
  } finally {
    clearTimeout(timer);
  }
}
