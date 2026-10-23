-- F-013: automatic reply drafts. Four tables (RLS on, no policies), RPC functions used only by the
-- Edge Functions with the secret key, and the pg_cron jobs. No subject, body, display name, AI output or token is stored.

create extension if not exists pg_cron;
create extension if not exists pg_net with schema extensions;

create table public.auto_draft_users (
  id uuid primary key default gen_random_uuid(),
  entra_oid text not null unique,
  mail_sha256 text not null,
  enabled boolean not null default false,
  status text not null default 'off' check (status in ('off', 'pending', 'active', 'not_in_scope', 'error')),
  enabled_changed_at timestamptz not null default now(),
  status_changed_at timestamptz not null default now(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Third-party personal data (chosen contacts). Plain lower-cased address only.
create table public.auto_reply_contacts (
  user_id uuid not null references public.auto_draft_users (id) on delete cascade,
  address text not null,
  created_at timestamptz not null default now(),
  primary key (user_id, address)
);

create table public.graph_subscriptions (
  user_id uuid primary key references public.auto_draft_users (id) on delete cascade,
  subscription_id text not null unique,
  expires_at timestamptz not null,
  last_renewed_at timestamptz not null default now(),
  last_catchup_at timestamptz
);

create table public.auto_draft_queue (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.auto_draft_users (id) on delete cascade,
  message_id text not null,
  status text not null default 'pending' check (status in ('pending', 'processing', 'drafted', 'skipped', 'failed')),
  reason_code text,
  attempts int not null default 0,
  next_attempt_at timestamptz not null default now(),
  locked_until timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, message_id)
);
create index auto_draft_queue_work on public.auto_draft_queue (status, next_attempt_at);
create index auto_draft_queue_user on public.auto_draft_queue (user_id, status);

alter table public.auto_draft_users enable row level security;
alter table public.auto_reply_contacts enable row level security;
alter table public.graph_subscriptions enable row level security;
alter table public.auto_draft_queue enable row level security;
-- Deliberately no policies: nothing is reachable through the Data API with the publishable key.
revoke all on public.auto_draft_users, public.auto_reply_contacts, public.graph_subscriptions, public.auto_draft_queue from anon, authenticated;

-- ---- RPC functions (called only by the Edge Functions) ----

create function public.ad_settings_get(p_oid text) returns jsonb language sql stable set search_path = '' as $$
  select coalesce(
    (select jsonb_build_object(
       'enabled', u.enabled,
       'status', u.status,
       'contacts', coalesce((select jsonb_agg(c.address order by c.address) from public.auto_reply_contacts c where c.user_id = u.id), '[]'::jsonb))
     from public.auto_draft_users u where u.entra_oid = p_oid),
    jsonb_build_object('enabled', false, 'status', 'off', 'contacts', '[]'::jsonb));
$$;

-- Replaces the list atomically (one function call is one transaction).
create function public.ad_settings_put(p_oid text, p_enabled boolean, p_sha text, p_contacts text[]) returns jsonb language plpgsql set search_path = '' as $$
declare
  v_id uuid; v_old boolean; v_status text; v_on boolean; v_new text;
begin
  insert into public.auto_draft_users (entra_oid, mail_sha256) values (p_oid, p_sha)
  on conflict (entra_oid) do update set mail_sha256 = excluded.mail_sha256
  returning id into v_id;
  select enabled, status into v_old, v_status from public.auto_draft_users where id = v_id for update;
  delete from public.auto_reply_contacts where user_id = v_id;
  insert into public.auto_reply_contacts (user_id, address) select v_id, a from unnest(coalesce(p_contacts, '{}')) as a;
  v_on := p_enabled and coalesce(cardinality(p_contacts), 0) > 0;
  v_new := case when not v_on then 'off' when v_old and v_status <> 'off' then v_status else 'pending' end;
  update public.auto_draft_users set
    enabled = p_enabled,
    status = v_new,
    status_changed_at = case when status is distinct from v_new then now() else status_changed_at end,
    enabled_changed_at = case when enabled is distinct from p_enabled then now() else enabled_changed_at end,
    updated_at = now()
  where id = v_id;
  return public.ad_settings_get(p_oid);
end $$;

create function public.ad_settings_delete(p_oid text) returns jsonb language plpgsql set search_path = '' as $$
begin
  delete from public.auto_reply_contacts where user_id = (select id from public.auto_draft_users where entra_oid = p_oid);
  update public.auto_draft_users set enabled = false, status = 'off', enabled_changed_at = now(), status_changed_at = now(), updated_at = now()
  where entra_oid = p_oid;
  return public.ad_settings_get(p_oid);
end $$;

create function public.ad_user_for_subscription(p_sub text) returns jsonb language sql stable set search_path = '' as $$
  select to_jsonb((select user_id from public.graph_subscriptions where subscription_id = p_sub));
$$;

create function public.ad_user_ctx(p_user uuid) returns jsonb language sql stable set search_path = '' as $$
  select to_jsonb(r) from (
    select u.entra_oid as oid, u.mail_sha256 as "mailSha256", u.enabled,
      coalesce((select jsonb_agg(c.address) from public.auto_reply_contacts c where c.user_id = u.id), '[]'::jsonb) as contacts
    from public.auto_draft_users u where u.id = p_user) r;
$$;

create function public.ad_cron_users() returns jsonb language sql stable set search_path = '' as $$
  select coalesce(jsonb_agg(to_jsonb(r)), '[]'::jsonb) from (
    select u.id, u.entra_oid as oid, u.status,
      (u.enabled and exists (select 1 from public.auto_reply_contacts c where c.user_id = u.id)) as active,
      s.subscription_id as "subscriptionId", s.expires_at as "expiresAt",
      s.last_renewed_at as "lastRenewedAt", s.last_catchup_at as "lastCatchupAt"
    from public.auto_draft_users u left join public.graph_subscriptions s on s.user_id = u.id
    where s.user_id is not null or (u.enabled and exists (select 1 from public.auto_reply_contacts c where c.user_id = u.id))
  ) r;
$$;

create function public.ad_sub_set(p_user uuid, p_sub text, p_expires timestamptz) returns void language sql set search_path = '' as $$
  insert into public.graph_subscriptions (user_id, subscription_id, expires_at, last_renewed_at)
  values (p_user, p_sub, p_expires, now())
  on conflict (user_id) do update set subscription_id = excluded.subscription_id, expires_at = excluded.expires_at, last_renewed_at = now();
$$;
create function public.ad_sub_delete(p_user uuid) returns void language sql set search_path = '' as $$
  delete from public.graph_subscriptions where user_id = p_user;
$$;
create function public.ad_catchup_done(p_user uuid, p_at timestamptz) returns void language sql set search_path = '' as $$
  update public.graph_subscriptions set last_catchup_at = p_at where user_id = p_user;
$$;
create function public.ad_user_status(p_user uuid, p_status text) returns void language sql set search_path = '' as $$
  update public.auto_draft_users set status = p_status, status_changed_at = now(), updated_at = now()
  where id = p_user and status is distinct from p_status;
$$;

-- Duplicate notifications collapse on the unique key.
create function public.ad_enqueue(p_user uuid, p_ids text[]) returns int language plpgsql set search_path = '' as $$
declare n int;
begin
  insert into public.auto_draft_queue (user_id, message_id) select p_user, i from unnest(p_ids) as i
  on conflict (user_id, message_id) do nothing;
  get diagnostics n = row_count;
  return n;
end $$;

-- Claims at most one row per mailbox (4-concurrent-requests limit) with a 5-minute lease.
-- Eligible: pending and due, or processing with an expired lease. The advisory lock is taken
-- before the live-lease check so two concurrent claimers cannot both pick the same mailbox.
create function public.ad_claim(p_limit int, p_user uuid default null) returns jsonb language plpgsql set search_path = '' as $$
declare
  r record; v_row record; v_out jsonb := '[]'::jsonb; n int := 0;
begin
  -- Attempt cap on reclaim: a row whose lease expired after 5 attempts (MAX_ATTEMPTS in processor.ts) fails instead of being retaken.
  update public.auto_draft_queue set status = 'failed', reason_code = 'lease_expired', locked_until = null, updated_at = now()
  where status = 'processing' and locked_until < now() and attempts >= 5
    and (p_user is null or user_id = p_user);
  for r in
    select distinct q.user_id from public.auto_draft_queue q
    where (p_user is null or q.user_id = p_user)
      and ((q.status = 'pending' and q.next_attempt_at <= now()) or (q.status = 'processing' and q.locked_until < now()))
  loop
    exit when n >= p_limit;
    if not pg_try_advisory_xact_lock(hashtextextended(r.user_id::text, 0)) then continue; end if;
    if exists (select 1 from public.auto_draft_queue where user_id = r.user_id and status = 'processing' and locked_until >= now()) then continue; end if;
    v_row := null;
    update public.auto_draft_queue set status = 'processing', locked_until = now() + interval '5 minutes', attempts = attempts + 1, updated_at = now()
    where id = (
      select id from public.auto_draft_queue
      where user_id = r.user_id and ((status = 'pending' and next_attempt_at <= now()) or (status = 'processing' and locked_until < now()))
      order by created_at for update skip locked limit 1)
    returning id, user_id, message_id, attempts, created_at into v_row;
    if v_row.id is not null then
      v_out := v_out || jsonb_build_array(jsonb_build_object('id', v_row.id, 'userId', v_row.user_id, 'messageId', v_row.message_id, 'attempts', v_row.attempts, 'createdAt', v_row.created_at));
      n := n + 1;
    end if;
  end loop;
  return v_out;
end $$;

create function public.ad_finish(p_id uuid, p_status text, p_reason text) returns void language sql set search_path = '' as $$
  update public.auto_draft_queue set status = p_status, reason_code = p_reason, locked_until = null, updated_at = now()
  where id = p_id and status = 'processing';
$$;
create function public.ad_retry(p_id uuid, p_reason text, p_next timestamptz) returns void language sql set search_path = '' as $$
  update public.auto_draft_queue set status = 'pending', reason_code = p_reason, next_attempt_at = p_next, locked_until = null, updated_at = now()
  where id = p_id and status = 'processing';
$$;

create function public.ad_counts_today(p_user uuid) returns jsonb language sql stable set search_path = '' as $$
  select jsonb_build_object(
    'user', count(*) filter (where user_id = p_user),
    'global', count(*))
  from public.auto_draft_queue
  -- rows that reached the AI step: drafted, or failed with an ai_* reason (ai_unavailable, ai_output_invalid)
  where (status = 'drafted' or (status = 'failed' and reason_code like 'ai\_%')) and updated_at >= date_trunc('day', now() at time zone 'UTC') at time zone 'UTC';
$$;

-- Queue rows: 30 days. Users disabled for 90 days or not_in_scope for 30 days go with their contacts.
create function public.ad_retention() returns void language sql set search_path = '' as $$
  delete from public.auto_draft_queue where created_at < now() - interval '30 days';
  delete from public.auto_draft_users
  where ((enabled = false and enabled_changed_at < now() - interval '90 days')
     or (status = 'not_in_scope' and status_changed_at < now() - interval '30 days'))
    and not exists (select 1 from public.graph_subscriptions g where g.user_id = public.auto_draft_users.id);
$$;

revoke all on function
  public.ad_settings_get(text), public.ad_settings_put(text, boolean, text, text[]), public.ad_settings_delete(text),
  public.ad_user_for_subscription(text), public.ad_user_ctx(uuid), public.ad_cron_users(),
  public.ad_sub_set(uuid, text, timestamptz), public.ad_sub_delete(uuid), public.ad_catchup_done(uuid, timestamptz),
  public.ad_user_status(uuid, text), public.ad_enqueue(uuid, text[]), public.ad_claim(int, uuid),
  public.ad_finish(uuid, text, text), public.ad_retry(uuid, text, timestamptz), public.ad_counts_today(uuid), public.ad_retention()
  from public, anon, authenticated;
grant execute on function
  public.ad_settings_get(text), public.ad_settings_put(text, boolean, text, text[]), public.ad_settings_delete(text),
  public.ad_user_for_subscription(text), public.ad_user_ctx(uuid), public.ad_cron_users(),
  public.ad_sub_set(uuid, text, timestamptz), public.ad_sub_delete(uuid), public.ad_catchup_done(uuid, timestamptz),
  public.ad_user_status(uuid, text), public.ad_enqueue(uuid, text[]), public.ad_claim(int, uuid),
  public.ad_finish(uuid, text, text), public.ad_retry(uuid, text, timestamptz), public.ad_counts_today(uuid), public.ad_retention()
  to service_role;

-- ---- Cron: calls auto-draft-cron with the secret read from Vault (secret name: auto_draft_cron_secret) ----
-- Operator, once (never in this file): vault.create_secret('<cron secret>', 'auto_draft_cron_secret') and vault.create_secret('https://<project-ref>.supabase.co', 'project_url').
select cron.schedule('auto-draft-process', '*/2 * * * *', $job$
  select net.http_post(
    url := (select decrypted_secret from vault.decrypted_secrets where name = 'project_url') || '/functions/v1/auto-draft-cron',
    headers := jsonb_build_object('Content-Type', 'application/json',
      'x-cron-secret', (select decrypted_secret from vault.decrypted_secrets where name = 'auto_draft_cron_secret')),
    body := '{"op":"process"}'::jsonb, timeout_milliseconds := 60000);
$job$);
select cron.schedule('auto-draft-renew', '0 */6 * * *', $job$
  select net.http_post(
    url := (select decrypted_secret from vault.decrypted_secrets where name = 'project_url') || '/functions/v1/auto-draft-cron',
    headers := jsonb_build_object('Content-Type', 'application/json',
      'x-cron-secret', (select decrypted_secret from vault.decrypted_secrets where name = 'auto_draft_cron_secret')),
    body := '{"op":"renew"}'::jsonb, timeout_milliseconds := 60000);
$job$);
select cron.schedule('auto-draft-retention', '17 3 * * *', $job$
  select net.http_post(
    url := (select decrypted_secret from vault.decrypted_secrets where name = 'project_url') || '/functions/v1/auto-draft-cron',
    headers := jsonb_build_object('Content-Type', 'application/json',
      'x-cron-secret', (select decrypted_secret from vault.decrypted_secrets where name = 'auto_draft_cron_secret')),
    body := '{"op":"retention"}'::jsonb, timeout_milliseconds := 60000);
$job$);
