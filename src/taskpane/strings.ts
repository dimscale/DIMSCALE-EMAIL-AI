// All UI strings live here (English only; localization is out of scope for F-001).
export const strings = {
  title: "DIMSCALE Mail AI",
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
};
