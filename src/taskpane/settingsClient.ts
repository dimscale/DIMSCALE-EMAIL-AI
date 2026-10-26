/* global fetch, AbortController, setTimeout, clearTimeout */
import { SETTINGS_URL } from "./config";

export type ServerStatus = "off" | "pending" | "active" | "not_in_scope" | "error";

export type SettingsData = {
  enabled: boolean;
  status: ServerStatus;
  contacts: string[];
};

export type SettingsResult =
  { kind: "ok"; data: SettingsData } | { kind: "rejected" } | { kind: "error"; status?: number };

export async function fetchSettings(
  token: string,
  fetchFn: typeof fetch = fetch
): Promise<SettingsResult> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 10000);
  try {
    const res = await fetchFn(SETTINGS_URL, {
      method: "GET",
      headers: { Authorization: `Bearer ${token}` },
      signal: ctrl.signal,
    });
    if (res.status === 401) return { kind: "rejected" };
    if (!res.ok) return { kind: "error", status: res.status };
    const body = await res.json();
    return {
      kind: "ok",
      data: {
        enabled: Boolean(body?.enabled),
        status: (body?.status as ServerStatus) || "off",
        contacts: Array.isArray(body?.contacts) ? body.contacts : [],
      },
    };
  } catch {
    return { kind: "error" };
  } finally {
    clearTimeout(timer);
  }
}

export async function saveSettings(
  token: string,
  payload: { enabled: boolean; ownAddress: string; contacts: string[] },
  fetchFn: typeof fetch = fetch
): Promise<SettingsResult> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 10000);
  try {
    const res = await fetchFn(SETTINGS_URL, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
      signal: ctrl.signal,
    });
    if (res.status === 401) return { kind: "rejected" };
    if (!res.ok) return { kind: "error", status: res.status };
    const body = await res.json();
    return {
      kind: "ok",
      data: {
        enabled: Boolean(body?.enabled),
        status: (body?.status as ServerStatus) || "off",
        contacts: Array.isArray(body?.contacts) ? body.contacts : [],
      },
    };
  } catch {
    return { kind: "error" };
  } finally {
    clearTimeout(timer);
  }
}

export async function deleteSettings(
  token: string,
  fetchFn: typeof fetch = fetch
): Promise<{ ok: boolean }> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 10000);
  try {
    const res = await fetchFn(SETTINGS_URL, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
      signal: ctrl.signal,
    });
    return { ok: res.ok };
  } catch {
    return { ok: false };
  } finally {
    clearTimeout(timer);
  }
}
