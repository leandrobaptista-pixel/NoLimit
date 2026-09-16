-- No Limit Administration: collaborative-session telemetry and safe workspace writes.
-- Apply in the No Limit Admin Supabase project's SQL editor before deploying app.js.
-- This migration is additive: it does not delete or rewrite existing operational records.

begin;

alter table public.user_presence
  add column if not exists device_platform text not null default 'Unknown',
  add column if not exists browser_name text not null default 'Unknown',
  add column if not exists viewport text not null default 'Unknown',
  add column if not exists last_sync_at timestamptz,
  add column if not exists last_sync_latency_ms integer;

-- The Admin currently stores its shared test workspace as one document.  This
-- function prevents a later browser from silently overwriting data that was
-- saved by somebody else after it loaded the document.  A conflicting browser
-- is told to refresh; it never replaces the newer cloud state.
create or replace function public.save_admin_workspace(
  workspace_org uuid,
  workspace_id text,
  workspace_payload jsonb,
  expected_updated_at timestamptz
)
returns table(updated_at timestamptz)
language plpgsql
security invoker
set search_path = public
as $$
begin
  update public.beta_workspaces
     set payload = workspace_payload,
         updated_by = auth.uid(),
         updated_at = clock_timestamp()
   where organization_id = workspace_org
     and id = workspace_id
     and updated_at = expected_updated_at
  returning beta_workspaces.updated_at into updated_at;

  if not found then
    raise exception 'The workspace changed on another device. Refresh before saving again.'
      using errcode = '40001';
  end if;

  return next;
end;
$$;

grant execute on function public.save_admin_workspace(uuid, text, jsonb, timestamptz) to authenticated;

commit;
