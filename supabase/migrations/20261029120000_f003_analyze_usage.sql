-- F-003: daily request counter for analyze-message. Only a SHA-256 hex of the lower-cased Entra oid, a UTC date
-- and a count are stored (no content, no AI output). Rows older than 2 days are deleted by an_take itself.

create table public.analyze_usage (
  user_key text not null,
  day date not null,
  count int not null default 0,
  primary key (user_key, day)
);

alter table public.analyze_usage enable row level security;
-- Deliberately no policies: nothing is reachable through the Data API with the publishable key.
revoke all on public.analyze_usage from anon, authenticated;

-- Takes one slot atomically: true when the user is under p_cap and today's total is under p_global.
-- One global advisory lock per transaction serialises callers, so the global sum check is not racy.
create function public.an_take(p_user text, p_cap int, p_global int) returns boolean language plpgsql set search_path = '' as $$
declare
  v_day date := (now() at time zone 'UTC')::date;
  v_user int; v_total int;
begin
  perform pg_advisory_xact_lock(hashtextextended('analyze_usage', 0));
  delete from public.analyze_usage where day < v_day - 2;
  select coalesce(sum(count), 0), coalesce(sum(count) filter (where user_key = p_user), 0)
    into v_total, v_user from public.analyze_usage where day = v_day;
  if v_user >= p_cap or v_total >= p_global then return false; end if;
  insert into public.analyze_usage (user_key, day, count) values (p_user, v_day, 1)
  on conflict (user_key, day) do update set count = public.analyze_usage.count + 1;
  return true;
end $$;

revoke all on function public.an_take(text, int, int) from public, anon, authenticated;
grant execute on function public.an_take(text, int, int) to service_role;
