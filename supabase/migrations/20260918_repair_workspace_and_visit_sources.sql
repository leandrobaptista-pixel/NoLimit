begin;
-- Qualify table columns: output parameter updated_at otherwise shadows the column.
create or replace function public.save_admin_workspace(
  workspace_org uuid, workspace_id text, workspace_payload jsonb,
  expected_updated_at timestamptz
) returns table(updated_at timestamptz)
language plpgsql set search_path = public as $$
begin
  update public.beta_workspaces as workspace
  set payload = workspace_payload, updated_by = auth.uid(), updated_at = clock_timestamp()
  where workspace.organization_id = workspace_org
    and workspace.id = workspace_id
    and workspace.updated_at = expected_updated_at
  returning workspace.updated_at into updated_at;
  if not found then
    raise exception 'The workspace changed on another device. Refresh before saving again.' using errcode = '40001';
  end if;
  return next;
end;
$$;
-- Preserve legacy numeric sources; new website submissions have UUID source IDs.
alter table public.admin_visit_requests alter column source_request_id drop not null;
alter table public.admin_visit_requests add column if not exists source_record_id uuid;
create unique index if not exists admin_visit_requests_source_record_idx
  on public.admin_visit_requests(source_record_id);
commit;
