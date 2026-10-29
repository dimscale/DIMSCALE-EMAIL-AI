// Non-secret config only (F-002/AC-16).
export const TENANT_ID = "c4333a40-3082-416f-80a2-dd502b43f30f"; // DIMSCALE tenant, from public OpenID metadata
export const CLIENT_ID = "d157a20c-8820-43ad-8dbf-174a5cade5ea"; // Entra app "DIMSCALE AI Email Assistant"
export const FUNCTION_URL = "https://qepzpkseyzclxnqguqta.supabase.co/functions/v1/whoami";
export const SETTINGS_URL =
  "https://qepzpkseyzclxnqguqta.supabase.co/functions/v1/auto-reply-settings";
export const ANALYZE_URL = "https://qepzpkseyzclxnqguqta.supabase.co/functions/v1/analyze-message";

export const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}`;
export const SCOPES = [`api://${CLIENT_ID}/access_as_user`, "profile"];

export const isConfigured = () =>
  !TENANT_ID.startsWith("00000000") &&
  !CLIENT_ID.startsWith("00000000") &&
  !FUNCTION_URL.includes("PROJECT-REF");
