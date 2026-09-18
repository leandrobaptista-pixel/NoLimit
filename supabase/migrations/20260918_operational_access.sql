begin;
grant select on public.beta_workspaces to service_role;
alter table public.organization_members add column if not exists linked_person_id text;

-- Build a deliberately small partner view. Financials and other people's records
-- never leave the server. Missing links result in an empty workspace.
create or replace function public.partner_workspace_projection(snapshot jsonb, person_id text)
returns jsonb language sql immutable set search_path=public as $$
with person as (
 select p from jsonb_array_elements(coalesce(snapshot->'people','[]'::jsonb)) p where p->>'id'=person_id
), ids as (
 select jsonb_array_elements_text(coalesce(p->'projectIds','[]'::jsonb)) id from person
), selected_projects as (
 select jsonb_build_object('id',p->>'id','name',p->>'name','status',p->>'status','service',p->>'service',
 'siteStreet',p->>'siteStreet','siteCity',p->>'siteCity','siteState',p->>'siteState','sitePostalCode',p->>'sitePostalCode',
 'startDate',p->>'startDate','progress',p->'progress','clientName','','clientId','') p
 from jsonb_array_elements(coalesce(snapshot->'projects','[]'::jsonb)) p where p->>'id' in(select id from ids)
)
select jsonb_build_object(
 'projects',coalesce((select jsonb_agg(p) from selected_projects),'[]'::jsonb),
 'people',coalesce((select jsonb_agg(p) from person),'[]'::jsonb),
 'schedule',coalesce((select jsonb_agg(p) from jsonb_array_elements(coalesce(snapshot->'schedule','[]'::jsonb)) p
 where p->>'personId'=person_id and p->>'projectId' in(select id from ids)),'[]'::jsonb),
 'media',coalesce((select jsonb_agg(p) from jsonb_array_elements(coalesce(snapshot->'media','[]'::jsonb)) p
 where p->>'projectId' in(select id from ids) and p->>'publishStatus' in ('approved','published')),'[]'::jsonb)
);
$$;
revoke all on function public.partner_workspace_projection(jsonb,text) from public;

create or replace function public.read_operational_workspace(workspace_org uuid, workspace_id text)
returns table(payload jsonb,updated_at timestamptz)
language plpgsql security definer set search_path=public as $$
declare member_role text; person_id text;
begin
 select m.role::text,m.linked_person_id into member_role,person_id from public.organization_members m
 where m.organization_id=workspace_org and m.user_id=auth.uid() and m.status='active';
 if member_role is null then raise exception 'Active membership required' using errcode='42501'; end if;
 return query select case when public.is_staff(workspace_org) then w.payload
 else jsonb_build_object('state',public.partner_workspace_projection(w.payload->'state',person_id)) end,w.updated_at
 from public.beta_workspaces w where w.organization_id=workspace_org and w.id=workspace_id;
end; $$;
revoke all on function public.read_operational_workspace(uuid,text) from public;
grant execute on function public.read_operational_workspace(uuid,text) to authenticated;

drop policy if exists beta_workspace_member_select on public.beta_workspaces;
drop policy if exists beta_workspace_member_insert on public.beta_workspaces;
drop policy if exists beta_workspace_member_update on public.beta_workspaces;
create policy beta_workspace_staff_select on public.beta_workspaces for select to authenticated using(public.is_staff(organization_id));
create policy beta_workspace_staff_insert on public.beta_workspaces for insert to authenticated with check(public.is_staff(organization_id) and updated_by=auth.uid());
create policy beta_workspace_staff_update on public.beta_workspaces for update to authenticated using(public.is_staff(organization_id)) with check(public.is_staff(organization_id) and updated_by=auth.uid());

-- Private files: staff, or the partner's own compliance documents / approved assigned media.
create or replace function public.can_read_workspace_file(bucket text, object_name text)
returns boolean language plpgsql stable security definer set search_path=public as $$
declare org uuid; person_id text; snapshot jsonb;
begin
 if bucket not in ('project-files','compliance-files','media-library') then return false; end if;
 begin org := split_part(object_name,'/',1)::uuid; exception when invalid_text_representation then return false; end;
 if public.is_staff(org) then return true; end if;
 select m.linked_person_id into person_id from public.organization_members m
 where m.organization_id=org and m.user_id=auth.uid() and m.status='active';
 if person_id is null then return false; end if;
 select public.partner_workspace_projection(w.payload->'state',person_id) into snapshot from public.beta_workspaces w
 where w.organization_id=org and w.id='shared-v1';
 if bucket='compliance-files' then
   return exists(select 1 from jsonb_array_elements(coalesce(snapshot->'people','[]'::jsonb)) p
   where p->>'w9FilePath'=object_name or p->>'insuranceFilePath'=object_name);
 elsif bucket='media-library' then
   return exists(select 1 from jsonb_array_elements(coalesce(snapshot->'media','[]'::jsonb)) p where p->>'filePath'=object_name);
 end if;
 return false;
end; $$;
revoke all on function public.can_read_workspace_file(text,text) from public;
grant execute on function public.can_read_workspace_file(text,text) to authenticated;
drop policy if exists nl_storage_read on storage.objects;
create policy nl_storage_read on storage.objects for select to authenticated using(public.can_read_workspace_file(bucket_id,name));
commit;
