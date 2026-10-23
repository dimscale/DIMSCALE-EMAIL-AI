export type Get = (k: string) => string | undefined;

export type Settings = {
  maxAgeHours: number;
  skipSenderDomains: string[];
  internalSenders: "allow" | "skip";
  internalDomains: string[];
  dailyCapPerUser: number;
  dailyCapGlobal: number;
  filterAutoResponseSuppress: boolean;
  filterFocusedOther: boolean;
  openaiModel: string;
  openaiMaxOutputTokens: number;
};

/** Kill switch: only the exact string "true" turns the feature on; read per call, so a secret change needs no redeploy. */
export const autoDraftsEnabled = (get: Get) => get("AUTO_DRAFTS_ENABLED") === "true";

/** A missing or invalid value falls back to the default; `defaulted` lists the env names for the `config_default` log. */
export function readSettings(get: Get): { settings: Settings; defaulted: string[] } {
  const defaulted: string[] = [];
  const pick = <T>(key: string, parse: (raw: string) => T | undefined, def: T): T => {
    const raw = get(key)?.trim();
    const v = raw ? parse(raw) : undefined;
    if (v === undefined) {
      defaulted.push(key);
      return def;
    }
    return v;
  };
  const posInt = (raw: string) => (/^\d+$/.test(raw) && Number(raw) > 0 ? Number(raw) : undefined);
  const bool = (raw: string) => (raw === "true" ? true : raw === "false" ? false : undefined);
  const list = (raw: string) => raw.split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  const settings: Settings = {
    maxAgeHours: pick("MAX_AGE_HOURS", posInt, 24),
    skipSenderDomains: pick("SKIP_SENDER_DOMAINS", list, []),
    internalSenders: pick<"allow" | "skip">("INTERNAL_SENDERS", (r) => (r === "allow" || r === "skip" ? r : undefined), "allow"),
    internalDomains: pick("INTERNAL_DOMAINS", list, []),
    dailyCapPerUser: pick("DAILY_CAP_PER_USER", posInt, 30),
    dailyCapGlobal: pick("DAILY_CAP_GLOBAL", posInt, 200),
    filterAutoResponseSuppress: pick("FILTER_AUTO_RESPONSE_SUPPRESS", bool, true),
    filterFocusedOther: pick("FILTER_FOCUSED_OTHER", bool, true),
    // No default model: D5 decides. An empty model makes the AI step fail closed (config_missing).
    openaiModel: get("OPENAI_MODEL")?.trim() ?? "",
    openaiMaxOutputTokens: pick("OPENAI_MAX_OUTPUT_TOKENS", posInt, 800),
  };
  return { settings, defaulted };
}
