// Mirrors isNoReplyLocalPart in supabase/functions/_shared/filters.ts (F-013/AC-33, AC-49 + amendment 2026-11-03). Keep the rule equal.
// NFKC, drop ignorable/format characters, lower-case, trim, cut at the first "+". Cyrillic look-alikes are not handled (accepted limit).
const NOREPLY = /^(noreply|no-reply|donotreply|do-not-reply|mailer-daemon|postmaster)$/;

export const isNoReplyLocalPart = (address: string): boolean => {
  const local = (address.split("@")[0] ?? "")
    .normalize("NFKC")
    .replace(/[\p{Default_Ignorable_Code_Point}\p{Cf}]/gu, "")
    .toLowerCase()
    .trim();
  return NOREPLY.test(local.split("+")[0]);
};
