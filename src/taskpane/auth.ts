import type { IPublicClientApplication } from "@azure/msal-browser";
import { AUTHORITY, CLIENT_ID, SCOPES } from "./config";

/* global Office */

export type TokenResult =
  { kind: "token"; token: string } | { kind: "interaction" } | { kind: "error"; code: string };

export const isNaaSupported = (): boolean =>
  Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");

const msal = () => import(/* webpackMode: "eager" */ "@azure/msal-browser");
let app: Promise<IPublicClientApplication> | undefined;
const getApp = () => {
  // No cacheLocation override: MSAL default (spec open question 3).
  // Lazy import (eager webpack mode: same bundle) keeps MSAL out of unit tests that never reach auth.
  if (!app) {
    const p = msal().then((m) =>
      m.createNestablePublicClientApplication({
        auth: { clientId: CLIENT_ID, authority: AUTHORITY },
      })
    );
    p.catch(() => (app = undefined)); // allow Retry to recreate after a failed init
    app = p;
  }
  return app;
};

// Never log the error or token: only the MSAL errorCode is surfaced.
const toResult = async (e: unknown): Promise<TokenResult> =>
  e instanceof (await msal()).InteractionRequiredAuthError
    ? { kind: "interaction" }
    : { kind: "error", code: (e as { errorCode?: string })?.errorCode || "unknown_error" };

export async function getTokenSilent(): Promise<TokenResult> {
  try {
    const r = await (await getApp()).acquireTokenSilent({ scopes: SCOPES });
    return { kind: "token", token: r.accessToken };
  } catch (e) {
    return await toResult(e);
  }
}

/** Call only from a user click (popup). */
export async function signIn(): Promise<TokenResult> {
  try {
    const r = await (await getApp()).acquireTokenPopup({ scopes: SCOPES });
    return { kind: "token", token: r.accessToken };
  } catch (e) {
    return { kind: "error", code: (e as { errorCode?: string })?.errorCode || "unknown_error" };
  }
}
