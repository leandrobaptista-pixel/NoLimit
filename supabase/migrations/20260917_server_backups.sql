-- No Limit Administration: server-side logical backup policy and scheduler.
--
-- This is an additive migration.  All backup records live in the No Limit
-- Supabase project; no browser or local-machine storage is used as an
-- operational source of truth.
--
-- Scope note: these are recoverable logical application backups.  They retain
-- the Admin workspace state, membership metadata and a media-object manifest.
-- Supabase Auth credentials and Storage file bytes remain managed by Supabase
-- and are intentionally never copied into JSON backup rows.

create extension if not exists pg_cron;

begin;

create table if not exists public.admin_backup_policies (
  organization_id uuid primary key references public.organizations(id) on delete cascade,
  enabled boolean not null default true,
  timezone text not null default 'America/New_York',
  incremental_interval interval not null default interval '1 hour',
  full_backup_local_time time not null default time '02:00',
  created_at timestamptz not null default clock_timestamp(),
  updated_at timestamptz not null default clock_timestamp(),
  updated_by uuid references public.profiles(id) on delete set null,
  check (incremental_interval >= interval '1 hour')
);

create table if not exists public.admin_backup_runs (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  backup_kind text not null check (backup_kind in ('incremental', 'full')),
  started_at timestamptz not null default clock_timestamp(),
  finished_at timestamptz,
  status text not null default 'running' check (status in ('running', 'completed', 'failed')),
  source_from timestamptz,
  source_to timestamptz not null default clock_timestamp(),
  change_count integer not null default 0,
  snapshot jsonb not null default '{}'::jsonb,
  error_message text,
  created_by uuid references public.profiles(id) on delete set null
);

create index if not exists admin_backup_runs_lookup_idx
  on public.admin_backup_runs (organization_id, backup_kind, finished_at desc);

alter table public.admin_backup_policies enable row level security;
alter table public.admin_backup_runs enable row level security;

drop policy if exists admin_backup_policies_owner_developer_read on public.admin_backup_policies;
create policy admin_backup_policies_owner_developer_read
on public.admin_backup_policies
for select to authenticated
using (public.is_owner_or_developer(organization_id));

drop policy if exists admin_backup_runs_owner_developer_read on public.admin_backup_runs;
create policy admin_backup_runs_owner_developer_read
on public.admin_backup_runs
for select to authenticated
using (public.is_owner_or_developer(organization_id));

-- The only browser-callable write path.  It does not expose a direct table
-- write, and it permits scheduling changes only to Owner/Developer accounts.
create or replace function public.configure_admin_backup_policy(
  target_organization_id uuid,
  target_timezone text,
  target_full_backup_local_time time,
  target_incremental_interval interval default interval '1 hour'
)
returns public.admin_backup_policies
language plpgsql
security definer
set search_path = public
as $$
declare
  configured public.admin_backup_policies;
begin
  if not public.is_owner_or_developer(target_organization_id) then
    raise exception 'Only an Owner or Developer can configure backups.' using errcode = '42501';
  end if;

  if target_incremental_interval < interval '1 hour' then
    raise exception 'Incremental backups cannot run more frequently than once per hour.' using errcode = '22023';
  end if;

  -- Raises if PostgreSQL does not recognize the IANA timezone name.
  perform now() at time zone target_timezone;

  insert into public.admin_backup_policies (
    organization_id, timezone, full_backup_local_time, incremental_interval, updated_by
  ) values (
    target_organization_id, target_timezone, target_full_backup_local_time,
    target_incremental_interval, auth.uid()
  )
  on conflict (organization_id) do update
    set enabled = true,
        timezone = excluded.timezone,
        full_backup_local_time = excluded.full_backup_local_time,
        incremental_interval = excluded.incremental_interval,
        updated_at = clock_timestamp(),
        updated_by = auth.uid()
  returning * into configured;

  insert into public.system_audit_events (
    organization_id, actor_user_id, action, department, entity_type, entity_id, metadata
  ) values (
    target_organization_id, auth.uid(), 'backup.policy_configured', 'security', 'backup_policy',
    target_organization_id::text,
    jsonb_build_object('timezone', target_timezone, 'full_backup_local_time', target_full_backup_local_time,
                       'incremental_interval', target_incremental_interval::text)
  );

  return configured;
end;
$$;

-- Kept private: it is executed only by the scheduled server runner below.
create or replace function public.create_admin_backup(
  target_organization_id uuid,
  requested_kind text,
  range_start timestamptz default null
)
returns uuid
language plpgsql
security definer
set search_path = public, storage
as $$
declare
  run_id uuid;
  range_end timestamptz := clock_timestamp();
  content jsonb;
  changed_rows integer := 0;
begin
  if requested_kind not in ('incremental', 'full') then
    raise exception 'Unsupported backup kind.' using errcode = '22023';
  end if;

  insert into public.admin_backup_runs (
    organization_id, backup_kind, source_from, source_to, created_by
  ) values (
    target_organization_id, requested_kind, range_start, range_end, auth.uid()
  ) returning id into run_id;

  if requested_kind = 'incremental' then
    select coalesce(jsonb_agg(jsonb_build_object(
      'workspace_id', version.workspace_id,
      'source_updated_at', version.source_updated_at,
      'saved_at', version.saved_at,
      'change_kind', version.change_kind,
      'payload', version.payload
    ) order by version.source_updated_at), '[]'::jsonb), count(*)
    into content, changed_rows
    from public.admin_workspace_versions version
    where version.organization_id = target_organization_id
      and (range_start is null or version.source_updated_at > range_start)
      and version.source_updated_at <= range_end;

    content := jsonb_build_object(
      'format', 'nolimit-admin/incremental-v1',
      'captured_at', range_end,
      'from', range_start,
      'to', range_end,
      'workspace_versions', content
    );
  else
    select jsonb_build_object(
      'format', 'nolimit-admin/full-v1',
      'captured_at', range_end,
      'workspaces', coalesce((
        select jsonb_agg(jsonb_build_object(
          'workspace_id', workspace.id,
          'updated_at', workspace.updated_at,
          'updated_by', workspace.updated_by,
          'payload', workspace.payload
        ) order by workspace.id)
        from public.beta_workspaces workspace
        where workspace.organization_id = target_organization_id
      ), '[]'::jsonb),
      'members', coalesce((
        select jsonb_agg(jsonb_build_object(
          'user_id', member.user_id,
          'role', member.role,
          'status', member.status
        ) order by member.user_id)
        from public.organization_members member
        where member.organization_id = target_organization_id
      ), '[]'::jsonb),
      'media_manifest', coalesce((
        select jsonb_agg(jsonb_build_object(
          'bucket_id', storage_object.bucket_id,
          'name', storage_object.name,
          'id', storage_object.id,
          'created_at', storage_object.created_at,
          'updated_at', storage_object.updated_at,
          'metadata', storage_object.metadata
        ) order by storage_object.name)
        from storage.objects storage_object
        where storage_object.name like target_organization_id::text || '/%'
      ), '[]'::jsonb)
    ) into content;

    select count(*) into changed_rows
    from public.beta_workspaces workspace
    where workspace.organization_id = target_organization_id;
  end if;

  update public.admin_backup_runs
     set status = 'completed',
         finished_at = clock_timestamp(),
         change_count = changed_rows,
         snapshot = content
   where id = run_id;

  insert into public.system_audit_events (
    organization_id, actor_user_id, action, department, entity_type, entity_id, metadata
  ) values (
    target_organization_id, auth.uid(), 'backup.completed', 'security', 'backup_run', run_id::text,
    jsonb_build_object('kind', requested_kind, 'change_count', changed_rows,
                       'source_from', range_start, 'source_to', range_end)
  );

  return run_id;
exception when others then
  update public.admin_backup_runs
     set status = 'failed', finished_at = clock_timestamp(), error_message = sqlerrm
   where id = run_id;
  raise;
end;
$$;

create or replace function public.run_due_admin_backups()
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  policy public.admin_backup_policies%rowtype;
  local_now timestamp;
  latest_backup timestamptz;
  full_already_ran boolean;
begin
  if not pg_try_advisory_xact_lock(hashtext('nolimit-admin-backup-runner')) then
    return;
  end if;

  for policy in select * from public.admin_backup_policies where enabled loop
    local_now := clock_timestamp() at time zone policy.timezone;
    select max(finished_at) into latest_backup
      from public.admin_backup_runs
     where organization_id = policy.organization_id
       and status = 'completed';

    select exists(
      select 1 from public.admin_backup_runs
       where organization_id = policy.organization_id
         and backup_kind = 'full'
         and status = 'completed'
         and (finished_at at time zone policy.timezone)::date = local_now::date
    ) into full_already_ran;

    if local_now::time >= policy.full_backup_local_time and not full_already_ran then
      perform public.create_admin_backup(policy.organization_id, 'full', latest_backup);
    elsif latest_backup is null or clock_timestamp() - latest_backup >= policy.incremental_interval then
      perform public.create_admin_backup(policy.organization_id, 'incremental', latest_backup);
    end if;
  end loop;
end;
$$;

revoke all on function public.create_admin_backup(uuid, text, timestamptz) from public, anon, authenticated;
revoke all on function public.run_due_admin_backups() from public, anon, authenticated;
revoke all on function public.configure_admin_backup_policy(uuid, text, time, interval) from public;
grant execute on function public.configure_admin_backup_policy(uuid, text, time, interval) to authenticated;

-- Production default: hourly incrementals and a complete snapshot at 2:00 AM
-- America/New_York. Owner/Developer may replace this policy through the secure
-- RPC above; the scheduler simply reads the current server policy.
insert into public.admin_backup_policies (organization_id, timezone, full_backup_local_time)
values ('00000000-0000-4000-8000-000000000001', 'America/New_York', time '02:00')
on conflict (organization_id) do nothing;

commit;

-- Check every five minutes; the policy function itself enforces the hourly
-- cadence and one complete backup for each local calendar day.
select cron.unschedule(jobid)
from cron.job
where jobname = 'nolimit-admin-backup-runner';

select cron.schedule(
  'nolimit-admin-backup-runner',
  '*/5 * * * *',
  $job$select public.run_due_admin_backups();$job$
);
