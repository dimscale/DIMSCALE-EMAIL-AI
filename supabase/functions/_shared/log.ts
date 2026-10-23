/** Fixed metadata fields only (F-013/AC-26). Anything else passed in is dropped, never printed. */
export type LogFields = {
  fn: string;
  id?: string;
  method?: string;
  status?: number;
  reason?: string;
  user_id?: string;
  attempt?: number;
  count?: number;
  ms?: number;
};
const KEYS = ["fn", "id", "method", "status", "reason", "user_id", "attempt", "count", "ms"] as const;

export function log(f: LogFields): void {
  const out: Record<string, unknown> = {};
  for (const k of KEYS) if (f[k] !== undefined) out[k] = f[k];
  console.log(JSON.stringify(out));
}
