/* global fetch, Response, URL, AbortController, setTimeout, clearTimeout */
import { getApp } from "./auth";

export const GRAPH_SCOPES = ["https://graph.microsoft.com/Mail.ReadBasic"];

export type GraphAuthResult =
  { kind: "token"; token: string } | { kind: "needConsent" } | { kind: "denied"; code?: string };

export type CandidateContact = {
  address: string;
  name: string;
};

const NO_REPLY_LOCAL_PARTS = new Set([
  "noreply",
  "no-reply",
  "donotreply",
  "do-not-reply",
  "mailer-daemon",
  "postmaster",
]);

export function isNoReplySender(address: string): boolean {
  const atIndex = address.indexOf("@");
  if (atIndex <= 0) return true;
  const localPart = address.slice(0, atIndex).toLowerCase();
  return NO_REPLY_LOCAL_PARTS.has(localPart);
}

export function filterRecentSenders(
  rawMessages: Array<{ from?: { emailAddress?: { address?: string; name?: string } } }>,
  ownAddress?: string
): CandidateContact[] {
  const own = ownAddress ? ownAddress.trim().toLowerCase() : "";
  const seen = new Set<string>();
  const candidates: CandidateContact[] = [];

  for (const m of rawMessages) {
    const rawAddr = m.from?.emailAddress?.address?.trim();
    if (!rawAddr) continue;
    const lower = rawAddr.toLowerCase();
    if (own && lower === own) continue;
    if (isNoReplySender(lower)) continue;
    if (seen.has(lower)) continue;

    seen.add(lower);
    candidates.push({
      address: lower,
      name: m.from?.emailAddress?.name?.trim() || "",
    });

    if (candidates.length >= 50) break;
  }

  return candidates;
}

export async function getGraphTokenSilent(): Promise<GraphAuthResult> {
  try {
    const app = await getApp();
    const r = await app.acquireTokenSilent({ scopes: GRAPH_SCOPES });
    return { kind: "token", token: r.accessToken };
  } catch (e) {
    const msal = await import("@azure/msal-browser");
    if (e instanceof msal.InteractionRequiredAuthError) {
      return { kind: "needConsent" };
    }
    const code = (e as { errorCode?: string })?.errorCode || "unknown_error";
    return { kind: "denied", code };
  }
}

/** Call only from a user click. */
export async function requestGraphConsent(): Promise<GraphAuthResult> {
  try {
    const app = await getApp();
    const r = await app.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    return { kind: "token", token: r.accessToken };
  } catch (e) {
    const code = (e as { errorCode?: string })?.errorCode || "consent_failed";
    return { kind: "denied", code };
  }
}

export async function fetchRecentSenders(
  token: string,
  ownAddress?: string,
  fetchFn: typeof fetch = fetch
): Promise<CandidateContact[]> {
  const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
  let url: string | undefined =
    `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages` +
    `?$select=from&$filter=receivedDateTime ge ${thirtyDaysAgo}&$orderby=receivedDateTime desc&$top=100`;

  const allMessages: Array<{ from?: { emailAddress?: { address?: string; name?: string } } }> = [];
  let totalFetched = 0;
  let pageCount = 0;

  while (url && totalFetched < 300 && pageCount < 5) {
    pageCount++;
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 10000);
    let res: Response;
    try {
      res = await fetchFn(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
        },
        signal: ctrl.signal,
      });
    } finally {
      clearTimeout(timer);
    }

    if (!res.ok) {
      throw new Error(`Graph request failed with status ${res.status}`);
    }

    const data: any = await res.json();
    const items = Array.isArray(data?.value) ? data.value : [];
    if (items.length === 0) {
      break;
    }
    allMessages.push(...items);
    totalFetched += items.length;

    const nextLink =
      typeof data?.["@odata.nextLink"] === "string" ? data["@odata.nextLink"] : undefined;
    if (nextLink) {
      try {
        const parsed = new URL(nextLink);
        if (
          parsed.protocol === "https:" &&
          parsed.hostname.toLowerCase() === "graph.microsoft.com"
        ) {
          url = nextLink;
        } else {
          url = undefined;
        }
      } catch {
        url = undefined;
      }
    } else {
      url = undefined;
    }
  }

  return filterRecentSenders(allMessages, ownAddress);
}
