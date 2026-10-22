// Non-secret config only (F-002/AC-16). Replace the placeholders once the Entra app is registered (D4).
export const TENANT_ID = "00000000-0000-0000-0000-000000000000"; // TODO(D4)
export const CLIENT_ID = "00000000-0000-0000-0000-000000000000"; // TODO(D4)
export const FUNCTION_URL = "https://qepzpkseyzclxnqguqta.supabase.co/functions/v1/whoami";

export const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}`;
export const SCOPES = [`api://${CLIENT_ID}/access_as_user`, "profile"];

export const isConfigured = () =>
  !TENANT_ID.startsWith("00000000") &&
  !CLIENT_ID.startsWith("00000000") &&
  !FUNCTION_URL.includes("PROJECT-REF");
