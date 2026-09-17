-- No Limit Administration: production synchronization, immutable workspace
-- history, and owner/developer-only audit access.
--
-- This migration is additive. It does not delete, replace, or expose existing
-- operational data. Apply it in the No Limit Admin Supabase project before the
-- corresponding Admin release is deployed.

begin;

-- The UI legitimately writes its own presence row after a successful cloud
-- save. Missing table grants caused those telemetry writes to fail and made a
-- successful workspace save look like a failure in the browser.
grant select, insert, update on table public.user_presence to authenticated;

-- A technical recovery/audit capability is deliberately narrower than the
-- normal Administrator role. It is assigned only as an organization membership
-- role of `owner` or `developer`.
create or replace function public.is_owner_or_developer(org uuid)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.organization_members member
    where member.organization_id = org
      and member.user_id = auth.uid()
      and member.status = 'active'
      and member.role in ('owner', 'developer')
  );
$$;

revoke all on function public.is_owner_or_developer(uuid) from public;
grant execute on function public.is_owner_or_developer(uuid) to authenticated;

-- Each confirmed workspace write receives a retained snapshot. This gives the
-- recovery system a real rollback point while the Admin is still migrating away
-- from its legacy shared-workspace document to individual operational tables.
create table if not exists public.admin_workspace_versions (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  workspace_id text not null,
  source_updated_at timestamptz not null,
  saved_at timestamptz not null default clock_timestamp(),
  saved_by uuid references public.profiles(id) on delete set null,
  change_kind text not null check (change_kind in ('created', 'updated', 'restored')),
  payload jsonb not null,
  unique (organization_id, workspace_id, source_updated_at)
);

create index if not exists admin_workspace_versions_lookup_idx
  on public.admin_workspace_versions (organization_id, workspace_id, saved_at desc);

alter table public.admin_workspace_versions enable row level security;

drop policy if exists admin_workspace_versions_owner_developer_read on public.admin_workspace_versions;
create policy admin_workspace_versions_owner_developer_read
on public.admin_workspace_versions
for select
to authenticated
using (public.is_owner_or_developer(organization_id));

-- No browser may insert, edit, or remove history. It is created solely by this
-- trigger when a protected workspace write succeeds.
create or replace function public.capture_admin_workspace_version()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.admin_workspace_versions (
    organization_id, workspace_id, source_updated_at, saved_by, change_kind, payload
  ) values (
    new.organization_id,
    new.id,
    new.updated_at,
    coalesce(new.updated_by, auth.uid()),
    case when tg_op = 'INSERT' then 'created' else 'updated' end,
    new.payload
  ) on conflict (organization_id, workspace_id, source_updated_at) do nothing;
  return new;
end;
$$;

drop trigger if exists capture_admin_workspace_version on public.beta_workspaces;
create trigger capture_admin_workspace_version
after insert or update of payload on public.beta_workspaces
for each row execute function public.capture_admin_workspace_version();

-- A compact, queryable event trail for operational and deployment activity.
-- Snapshot data belongs in admin_workspace_versions; this table contains only
-- action metadata and never passwords, reset tokens, or email contents.
create table if not exists public.system_audit_events (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  actor_user_id uuid references public.profiles(id) on delete set null,
  action text not null,
  department text not null default 'administration',
  entity_type text not null,
  entity_id text,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default clock_timestamp()
);

create index if not exists system_audit_events_org_time_idx
  on public.system_audit_events (organization_id, created_at desc);
create index if not exists system_audit_events_entity_idx
  on public.system_audit_events (organization_id, entity_type, entity_id, created_at desc);

alter table public.system_audit_events enable row level security;

drop policy if exists system_audit_events_owner_developer_read on public.system_audit_events;
create policy system_audit_events_owner_developer_read
on public.system_audit_events
for select
to authenticated
using (public.is_owner_or_developer(organization_id));

-- Capture the old shared-workspace save flow immediately. New normalized
-- tables will receive equivalent per-record trigger events as they are added.
create or replace function public.record_workspace_audit_event()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.system_audit_events (
    organization_id, actor_user_id, action, department, entity_type, entity_id, metadata
  ) values (
    new.organization_id,
    coalesce(new.updated_by, auth.uid()),
    case when tg_op = 'INSERT' then 'workspace.created' else 'workspace.updated' end,
    'administration',
    'workspace',
    new.id,
    jsonb_build_object('version_at', new.updated_at, 'changed_by', coalesce(new.updated_by, auth.uid()))
  );
  return new;
end;
$$;

drop trigger if exists record_workspace_audit_event on public.beta_workspaces;
create trigger record_workspace_audit_event
after insert or update of payload on public.beta_workspaces
for each row execute function public.record_workspace_audit_event();

commit;
