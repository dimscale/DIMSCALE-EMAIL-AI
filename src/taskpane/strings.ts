// All UI strings live here (English only; localization is out of scope for F-001).
export const strings = {
  title: "DIMSCALE Mail AI · test",
  noItem: "Open a message to use DIMSCALE Mail AI",
  loading: "Loading message...",
  noSubject: "(no subject)",
  noBody: "This message has no text content",
  truncated: "(truncated)",
  error: "Could not read the message body. Please try again.",
  retry: "Retry",
  subject: "Subject",
  from: "From",
  date: "Date",
  body: "Body",
  bodyLength: (n: number) => `Body length: ${n} characters`,
  backendWorking: "Connecting to DIMSCALE backend…",
  connected: "Backend connected",
  connectedAs: (name: string) => `Backend connected as ${name}`,
  unsupported:
    "This Outlook version doesn't support sign-in for DIMSCALE Mail AI. Update Outlook or use Outlook on the web.",
  needSignIn: "Sign in to connect to the DIMSCALE backend",
  signIn: "Sign in",
  signInFailed: "Sign-in failed",
  signInError: "Could not get a sign-in token",
  rejected: "Backend rejected the sign-in",
  unreachable: "Backend unreachable",
  notConfigured: "DIMSCALE backend sign-in is not configured yet", // F-002 deviation: state for placeholder config
};
