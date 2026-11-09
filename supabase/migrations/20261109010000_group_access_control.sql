-- Production access control. No application records are deleted or recreated.
create table if not exists public.nl_access_owner(organization_id uuid primary key references public.organizations(id), user_id uuid not null references auth.users(id));
alter table public.nl_access_owner enable row level security;
revoke all on public.nl_access_owner from public,anon,authenticated;
insert into public.nl_access_owner values ('00000000-0000-4000-8000-000000000001','664fe48b-cb30-4e2d-90ba-93ffe17eee97') on conflict do nothing;
create table if not exists public.nl_access_groups(organization_id uuid not null references public.organizations(id), group_key text not null, label text not null, scope text not null check(scope in ('all','assigned')), permissions jsonb not null, version integer not null default 1, updated_at timestamptz not null default clock_timestamp(), updated_by uuid references auth.users(id), primary key(organization_id,group_key));
alter table public.nl_access_groups enable row level security;
revoke all on public.nl_access_groups from public,anon,authenticated;
insert into public.nl_access_groups(organization_id,group_key,label,scope,permissions) values ('00000000-0000-4000-8000-000000000001','admin','Administration','all','{"overview": {"view": true, "create": false, "edit": false, "delete": false}, "requests": {"view": true, "create": true, "edit": true, "delete": false}, "documents": {"view": true, "create": true, "edit": true, "delete": false}, "clients": {"view": true, "create": true, "edit": true, "delete": false}, "projects": {"view": true, "create": true, "edit": true, "delete": false}, "financial": {"view": true, "create": true, "edit": true, "delete": false}, "services": {"view": true, "create": true, "edit": true, "delete": false}, "materials": {"view": true, "create": true, "edit": true, "delete": false}, "team": {"view": true, "create": true, "edit": true, "delete": false}, "vendors": {"view": true, "create": true, "edit": true, "delete": false}, "schedule": {"view": true, "create": true, "edit": true, "delete": false}, "media": {"view": true, "create": true, "edit": true, "delete": false}, "map": {"view": true, "create": false, "edit": false, "delete": false}, "compliance": {"view": true, "create": true, "edit": true, "delete": false}, "reports": {"view": true, "create": false, "edit": false, "delete": false}}'::jsonb) on conflict do nothing;
insert into public.nl_access_groups(organization_id,group_key,label,scope,permissions) values ('00000000-0000-4000-8000-000000000001','office','Finance','all','{"overview": {"view": true, "create": false, "edit": false, "delete": false}, "requests": {"view": false, "create": false, "edit": false, "delete": false}, "documents": {"view": true, "create": true, "edit": true, "delete": false}, "clients": {"view": true, "create": true, "edit": true, "delete": false}, "projects": {"view": true, "create": false, "edit": false, "delete": false}, "financial": {"view": true, "create": true, "edit": true, "delete": false}, "services": {"view": true, "create": true, "edit": true, "delete": false}, "materials": {"view": true, "create": true, "edit": true, "delete": false}, "team": {"view": false, "create": false, "edit": false, "delete": false}, "vendors": {"view": true, "create": false, "edit": false, "delete": false}, "schedule": {"view": false, "create": false, "edit": false, "delete": false}, "media": {"view": false, "create": false, "edit": false, "delete": false}, "map": {"view": false, "create": false, "edit": false, "delete": false}, "compliance": {"view": false, "create": false, "edit": false, "delete": false}, "reports": {"view": true, "create": false, "edit": false, "delete": false}}'::jsonb) on conflict do nothing;
insert into public.nl_access_groups(organization_id,group_key,label,scope,permissions) values ('00000000-0000-4000-8000-000000000001','manager','Project managers','all','{"overview": {"view": true, "create": false, "edit": false, "delete": false}, "requests": {"view": false, "create": false, "edit": false, "delete": false}, "documents": {"view": false, "create": false, "edit": false, "delete": false}, "clients": {"view": true, "create": false, "edit": false, "delete": false}, "projects": {"view": true, "create": true, "edit": true, "delete": false}, "financial": {"view": false, "create": false, "edit": false, "delete": false}, "services": {"view": false, "create": false, "edit": false, "delete": false}, "materials": {"view": true, "create": true, "edit": true, "delete": false}, "team": {"view": true, "create": false, "edit": false, "delete": false}, "vendors": {"view": true, "create": false, "edit": false, "delete": false}, "schedule": {"view": true, "create": true, "edit": true, "delete": false}, "media": {"view": true, "create": true, "edit": true, "delete": false}, "map": {"view": true, "create": false, "edit": false, "delete": false}, "compliance": {"view": true, "create": true, "edit": true, "delete": false}, "reports": {"view": true, "create": false, "edit": false, "delete": false}}'::jsonb) on conflict do nothing;
insert into public.nl_access_groups(organization_id,group_key,label,scope,permissions) values ('00000000-0000-4000-8000-000000000001','team_member','Employees','assigned','{"overview": {"view": true, "create": false, "edit": false, "delete": false}, "requests": {"view": false, "create": false, "edit": false, "delete": false}, "documents": {"view": false, "create": false, "edit": false, "delete": false}, "clients": {"view": false, "create": false, "edit": false, "delete": false}, "projects": {"view": true, "create": false, "edit": false, "delete": false}, "financial": {"view": false, "create": false, "edit": false, "delete": false}, "services": {"view": false, "create": false, "edit": false, "delete": false}, "materials": {"view": false, "create": false, "edit": false, "delete": false}, "team": {"view": true, "create": false, "edit": false, "delete": false}, "vendors": {"view": false, "create": false, "edit": false, "delete": false}, "schedule": {"view": true, "create": false, "edit": false, "delete": false}, "media": {"view": true, "create": false, "edit": false, "delete": false}, "map": {"view": true, "create": false, "edit": false, "delete": false}, "compliance": {"view": false, "create": false, "edit": false, "delete": false}, "reports": {"view": false, "create": false, "edit": false, "delete": false}}'::jsonb) on conflict do nothing;
insert into public.nl_access_groups(organization_id,group_key,label,scope,permissions) values ('00000000-0000-4000-8000-000000000001','vendor','Vendors','assigned','{"overview": {"view": true, "create": false, "edit": false, "delete": false}, "requests": {"view": false, "create": false, "edit": false, "delete": false}, "documents": {"view": false, "create": false, "edit": false, "delete": false}, "clients": {"view": false, "create": false, "edit": false, "delete": false}, "projects": {"view": true, "create": false, "edit": false, "delete": false}, "financial": {"view": false, "create": false, "edit": false, "delete": false}, "services": {"view": false, "create": false, "edit": false, "delete": false}, "materials": {"view": false, "create": false, "edit": false, "delete": false}, "team": {"view": false, "create": false, "edit": false, "delete": false}, "vendors": {"view": true, "create": false, "edit": false, "delete": false}, "schedule": {"view": true, "create": false, "edit": false, "delete": false}, "media": {"view": true, "create": false, "edit": false, "delete": false}, "map": {"view": true, "create": false, "edit": false, "delete": false}, "compliance": {"view": false, "create": false, "edit": false, "delete": false}, "reports": {"view": false, "create": false, "edit": false, "delete": false}}'::jsonb) on conflict do nothing;
insert into public.nl_access_groups(organization_id,group_key,label,scope,permissions) values ('00000000-0000-4000-8000-000000000001','subcontractor','Subcontractors','assigned','{"overview": {"view": true, "create": false, "edit": false, "delete": false}, "requests": {"view": false, "create": false, "edit": false, "delete": false}, "documents": {"view": false, "create": false, "edit": false, "delete": false}, "clients": {"view": false, "create": false, "edit": false, "delete": false}, "projects": {"view": true, "create": false, "edit": false, "delete": false}, "financial": {"view": false, "create": false, "edit": false, "delete": false}, "services": {"view": false, "create": false, "edit": false, "delete": false}, "materials": {"view": false, "create": false, "edit": false, "delete": false}, "team": {"view": true, "create": false, "edit": false, "delete": false}, "vendors": {"view": false, "create": false, "edit": false, "delete": false}, "schedule": {"view": true, "create": false, "edit": false, "delete": false}, "media": {"view": true, "create": false, "edit": false, "delete": false}, "map": {"view": true, "create": false, "edit": false, "delete": false}, "compliance": {"view": false, "create": false, "edit": false, "delete": false}, "reports": {"view": false, "create": false, "edit": false, "delete": false}}'::jsonb) on conflict do nothing;
insert into public.nl_access_groups(organization_id,group_key,label,scope,permissions) values ('00000000-0000-4000-8000-000000000001','viewer','Read only','assigned','{"overview": {"view": true, "create": false, "edit": false, "delete": false}, "requests": {"view": false, "create": false, "edit": false, "delete": false}, "documents": {"view": false, "create": false, "edit": false, "delete": false}, "clients": {"view": false, "create": false, "edit": false, "delete": false}, "projects": {"view": true, "create": false, "edit": false, "delete": false}, "financial": {"view": false, "create": false, "edit": false, "delete": false}, "services": {"view": false, "create": false, "edit": false, "delete": false}, "materials": {"view": false, "create": false, "edit": false, "delete": false}, "team": {"view": false, "create": false, "edit": false, "delete": false}, "vendors": {"view": false, "create": false, "edit": false, "delete": false}, "schedule": {"view": true, "create": false, "edit": false, "delete": false}, "media": {"view": true, "create": false, "edit": false, "delete": false}, "map": {"view": true, "create": false, "edit": false, "delete": false}, "compliance": {"view": false, "create": false, "edit": false, "delete": false}, "reports": {"view": false, "create": false, "edit": false, "delete": false}}'::jsonb) on conflict do nothing;
insert into public.nl_access_groups(organization_id,group_key,label,scope,permissions) values ('00000000-0000-4000-8000-000000000001','developer','Developer — owner only','assigned','{"overview": {"view": false, "create": false, "edit": false, "delete": false}, "requests": {"view": false, "create": false, "edit": false, "delete": false}, "documents": {"view": false, "create": false, "edit": false, "delete": false}, "clients": {"view": false, "create": false, "edit": false, "delete": false}, "projects": {"view": false, "create": false, "edit": false, "delete": false}, "financial": {"view": false, "create": false, "edit": false, "delete": false}, "services": {"view": false, "create": false, "edit": false, "delete": false}, "materials": {"view": false, "create": false, "edit": false, "delete": false}, "team": {"view": false, "create": false, "edit": false, "delete": false}, "vendors": {"view": false, "create": false, "edit": false, "delete": false}, "schedule": {"view": false, "create": false, "edit": false, "delete": false}, "media": {"view": false, "create": false, "edit": false, "delete": false}, "map": {"view": false, "create": false, "edit": false, "delete": false}, "compliance": {"view": false, "create": false, "edit": false, "delete": false}, "reports": {"view": false, "create": false, "edit": false, "delete": false}}'::jsonb) on conflict do nothing;

create or replace function public.nl_is_owner(org uuid) returns boolean language sql stable security definer set search_path=public as $$
select exists(select 1 from nl_access_owner o join organization_members m on m.organization_id=o.organization_id and m.user_id=o.user_id where o.organization_id=org and o.user_id=auth.uid() and m.status='active' and m.role='owner'); $$;
create or replace function public.nl_permission_for(org uuid, actor uuid, area text, operation text) returns boolean language sql stable security definer set search_path=public as $$
select exists(select 1 from organization_members m left join nl_access_groups g on g.organization_id=m.organization_id and g.group_key=m.role::text where m.organization_id=org and m.user_id=actor and m.status='active' and (
 exists(select 1 from nl_access_owner o where o.organization_id=org and o.user_id=actor and m.role='owner')
 or (m.role not in ('owner','developer') and coalesce((g.permissions->area->>'view')::boolean,false) and coalesce((g.permissions->area->>operation)::boolean,false)))); $$;
revoke all on function public.nl_permission_for(uuid,uuid,text,text) from public,anon,authenticated;
grant execute on function public.nl_permission_for(uuid,uuid,text,text) to service_role;
create or replace function public.nl_has_permission(org uuid, area text, operation text) returns boolean language sql stable security definer set search_path=public as $$ select public.nl_permission_for(org,auth.uid(),area,operation); $$;
create or replace function public.nl_access_context(org uuid) returns jsonb language plpgsql stable security definer set search_path=public as $$
declare member organization_members; g nl_access_groups; result jsonb;
begin
 select * into member from organization_members where organization_id=org and user_id=auth.uid() and status='active';
 if not found then raise exception 'Active membership required.' using errcode='42501';end if;
 select * into g from nl_access_groups where organization_id=org and group_key=member.role::text;
 result:=jsonb_build_object('owner',nl_is_owner(org),'role',member.role,'scope',case when nl_is_owner(org) then 'all' else coalesce(g.scope,'assigned') end,'permissions',coalesce(g.permissions,'{}'),'version',coalesce(g.version,0),'linkedPersonId',member.linked_person_id);
 if nl_is_owner(org) then result:=result||jsonb_build_object('groups',(select jsonb_agg(to_jsonb(x)-'organization_id'-'updated_by' order by x.group_key) from nl_access_groups x where organization_id=org),'members',(select coalesce(jsonb_agg(jsonb_build_object('id',m.user_id,'name',coalesce(nullif(p.full_name,''),p.email),'email',p.email,'role',m.role,'status',m.status,'linkedPersonId',m.linked_person_id) order by p.email),'[]') from organization_members m left join profiles p on p.id=m.user_id where m.organization_id=org)); end if;
 return result;
end; $$;
create or replace function public.nl_save_access_group(org uuid, target_group text, target_scope text, target_permissions jsonb, expected_version integer) returns jsonb language plpgsql security definer set search_path=public as $$
declare area text; action text; old_group nl_access_groups; b boolean;
begin
 if not nl_is_owner(org) then raise exception 'Only the Owner can change permissions.' using errcode='42501';end if;
 if target_group not in ('admin','office','manager','team_member','vendor','subcontractor','viewer') or target_scope not in ('all','assigned') or jsonb_typeof(target_permissions)<>'object' then raise exception 'Invalid access group.';end if;
 if target_scope='assigned' and coalesce((target_permissions->'requests'->>'view')::boolean,false) then raise exception 'Website visit requests require all-company project scope.';end if;
 select * into strict old_group from nl_access_groups where organization_id=org and group_key=target_group for update;
 if old_group.version<>expected_version then raise exception 'Permissions changed on another device. Reload before saving.' using errcode='40001';end if;
 if (select count(*) from jsonb_object_keys(target_permissions))<>15 then raise exception 'Every department must be included.';end if;
 for area in select jsonb_object_keys(target_permissions) loop
 if area not in ('overview','requests','documents','clients','projects','financial','services','materials','team','vendors','schedule','media','map','compliance','reports') or jsonb_typeof(target_permissions->area)<>'object' or (select count(*) from jsonb_object_keys(target_permissions->area))<>4 then raise exception 'Invalid department.';end if;
 for action in select jsonb_object_keys(target_permissions->area) loop
 if action not in ('view','create','edit','delete') or jsonb_typeof(target_permissions->area->action)<>'boolean' then raise exception 'Invalid permission.';end if;
 end loop;
 if not (target_permissions->area->>'view')::boolean and ((target_permissions->area->>'create')::boolean or (target_permissions->area->>'edit')::boolean or (target_permissions->area->>'delete')::boolean) then raise exception 'Enable View before granting changes.';end if;
 if area in ('overview','map','reports') and ((target_permissions->area->>'create')::boolean or (target_permissions->area->>'edit')::boolean or (target_permissions->area->>'delete')::boolean) then raise exception 'This department is read only.';end if;
 end loop;
 update nl_access_groups set scope=target_scope, permissions=target_permissions,version=version+1,updated_at=clock_timestamp(),updated_by=auth.uid() where organization_id=org and group_key=target_group;
 insert into system_audit_events(organization_id,actor_user_id,action,department,entity_type,entity_id,metadata) values(org,auth.uid(),'access.permissions_changed','security','access_group',target_group,jsonb_build_object('before',jsonb_build_object('scope',old_group.scope,'permissions',old_group.permissions),'after',jsonb_build_object('scope',target_scope,'permissions',target_permissions)));
 return nl_access_context(org);
end; $$;
create or replace function public.nl_set_member_access(org uuid, target_user uuid, target_group text, target_status text, linked_person text default null) returns jsonb language plpgsql security definer set search_path=public as $$
declare previous organization_members; p jsonb;
begin
 if not nl_is_owner(org) then raise exception 'Only the Owner can change user access.' using errcode='42501';end if;
 if exists(select 1 from nl_access_owner where organization_id=org and user_id=target_user) then raise exception 'The Owner account is protected.' using errcode='42501';end if;
 if target_group not in ('admin','office','manager','team_member','vendor','subcontractor','viewer') or target_status not in ('active','inactive') then raise exception 'Invalid group or status.';end if;
 select * into strict previous from organization_members where organization_id=org and user_id=target_user for update;
 if nullif(linked_person,'') is not null then
 select r into p from beta_workspaces w,jsonb_array_elements(w.payload->'state'->'people') r where w.organization_id=org and w.id='shared-v1' and r->>'id'=linked_person;
 if p is null then raise exception 'Choose a registered person or company.';end if;
 end if;
 if target_status='active' and (target_group in ('vendor','subcontractor','team_member') or (select scope='assigned' from nl_access_groups where organization_id=org and group_key=target_group)) and p is null then raise exception 'Link a person or company before activating assigned-project access.';end if;
 if target_group='vendor' and p->>'type'<>'Vendor' then raise exception 'Choose a Vendor record.';end if;
 if target_group='subcontractor' and p->>'type'<>'Subcontractor' then raise exception 'Choose a Subcontractor record.';end if;
 update organization_members set role=target_group::app_role,status=target_status,linked_person_id=nullif(linked_person,'') where organization_id=org and user_id=target_user;
 insert into system_audit_events(organization_id,actor_user_id,action,department,entity_type,entity_id,metadata) values(org,auth.uid(),'access.member_changed','security','member',target_user::text,jsonb_build_object('previous_role',previous.role,'role',target_group,'status',target_status,'linked_person_id',linked_person));
 return nl_access_context(org);
end; $$;
-- A collection's owning department is checked again on every server write.
create or replace function public.nl_collection_area(collection text, item jsonb default '{}') returns text language sql immutable set search_path=public as $$ select case collection
 when 'people' then case when item->>'type'='Vendor' then 'vendors' else 'team' end
 when 'projects' then 'projects' when 'clients' then 'clients' when 'requests' then 'requests' when 'siteVisits' then 'requests'
 when 'estimates' then 'documents' when 'contracts' then 'documents' when 'invoices' then 'documents' when 'changeOrders' then 'documents'
 when 'materials' then 'materials' when 'customServices' then 'services'
 when 'payments' then 'financial' when 'transactions' then 'financial' when 'expenseReceipts' then 'financial' when 'projectEngagements' then 'financial' when 'providerPayments' then 'financial' when 'providerMilestones' then 'financial'
 when 'schedule' then 'schedule' when 'attendance' then 'schedule' when 'media' then 'media'
 when 'compliance' then 'compliance' when 'consents' then 'compliance' when 'insuranceRenewals' then 'compliance'
 else 'security' end; $$;
create or replace function public.nl_record_in_scope(org uuid, snapshot jsonb, collection text, item jsonb) returns boolean language plpgsql stable security definer set search_path=public as $$
declare person text; scope_value text; ids jsonb;
begin
 if nl_is_owner(org) then return true;end if;
 select m.linked_person_id,g.scope into person,scope_value from organization_members m join nl_access_groups g on g.organization_id=m.organization_id and g.group_key=m.role::text where m.organization_id=org and m.user_id=auth.uid() and m.status='active';
 if scope_value='all' then return true;end if;
 if person is null then return false;end if;
 select coalesce(p->'projectIds','[]') into ids from jsonb_array_elements(coalesce(snapshot->'people','[]')) p where p->>'id'=person;
 ids:=coalesce(ids,'[]');
 if collection='customServices' then return true;
 elsif collection='people' then return item->>'id'=person;
 elsif collection='projects' then return ids ? (item->>'id');
 elsif collection='clients' then return exists(select 1 from jsonb_array_elements(coalesce(snapshot->'projects','[]')) p where ids ? (p->>'id') and p->>'clientId'=item->>'id');
 elsif collection in ('schedule','attendance') then return item->>'personId'=person and ids ? (item->>'projectId');
 elsif collection='media' then return ids ? (item->>'projectId') and (item->>'publishStatus' in ('approved','published') or item->>'uploadedBy'=auth.uid()::text);
 elsif collection='insuranceRenewals' then return item->>'personId'=person;
 else return coalesce(ids ? (item->>'projectId'),false);
 end if;
end; $$;
create or replace function public.nl_visible_record(org uuid, collection text, item jsonb) returns jsonb language plpgsql stable security definer set search_path=public as $$
declare result jsonb:=item;
begin
 if nl_is_owner(org) then return item;end if;
 if not nl_has_permission(org,'financial','view') then
 if collection='projects' then result:=result-array['contractValue','cost','outstanding'];
 elsif collection='schedule' then select coalesce(jsonb_object_agg(key,value),'{}') into result from jsonb_each(item) where key=any(array['id','shift','status','personId','workDate','projectId','personName','personType','projectName','instructions','engagementId','createdAt','updatedAt','updatedBy']);
 elsif collection='people' then result:=result-array['dailyRate','hourlyRate','rate','balance','totalPaid','payRate'];end if;
 end if;
 if collection='projects' and not nl_has_permission(org,'documents','view') then result:=result-'materialTakeoffs';end if;
 if collection='projects' and not nl_has_permission(org,'clients','view') then result:=result-array['clientId','clientName'];end if;
 if collection='people' and not nl_has_permission(org,'compliance','view') then result:=result-array['w9FilePath','w9FileName','insuranceFilePath','insuranceFileName','policyNumber','coverageAmount'];end if;
 return result;
end; $$;
create or replace function public.nl_project_state(org uuid,snapshot jsonb) returns jsonb language plpgsql stable security definer set search_path=public as $$
declare result jsonb:='{}'; col text; rows jsonb; item jsonb; area text; allowed boolean;
begin
 if not is_member(org) then raise exception 'Active membership required.' using errcode='42501';end if;
 if nl_is_owner(org) then return snapshot;end if;
 for col,rows in select key,value from jsonb_each(snapshot) loop
 if jsonb_typeof(rows)<>'array' then continue;end if;
 result:=result||jsonb_build_object(col,'[]'::jsonb);
 for item in select value from jsonb_array_elements(rows) loop
 area:=nl_collection_area(col,item);
 allowed:=nl_has_permission(org,area,'view');
 if col='projects' then allowed:=allowed or nl_has_permission(org,'map','view');end if;
 if allowed and nl_record_in_scope(org,snapshot,col,item) then result:=jsonb_set(result,array[col],result->col||jsonb_build_array(nl_visible_record(org,col,item)));end if;
 end loop;
 end loop;
 return result;
end; $$;
-- Existing helpers now reserve raw legacy tables, security and recovery to the pinned Owner.
create or replace function public.is_staff(org uuid) returns boolean language sql stable security definer set search_path=public as $$ select nl_is_owner(org); $$;
create or replace function public.is_org_admin(org uuid) returns boolean language sql stable security definer set search_path=public as $$ select nl_is_owner(org); $$;
create or replace function public.is_owner_or_developer(org uuid) returns boolean language sql stable security definer set search_path=public as $$ select nl_is_owner(org); $$;
create or replace function public.read_operational_workspace(workspace_org uuid,workspace_id text) returns table(payload jsonb,updated_at timestamptz) language plpgsql stable security definer set search_path=public as $$
begin
 if not is_member(workspace_org) then raise exception 'Active membership required.' using errcode='42501';end if;
 return query select jsonb_build_object('state',nl_project_state(workspace_org,w.payload->'state')),w.updated_at from beta_workspaces w where w.organization_id=workspace_org and w.id=workspace_id;
end; $$;
create or replace function public.save_admin_workspace(workspace_org uuid,workspace_id text,workspace_payload jsonb,expected_updated_at timestamptz) returns table(updated_at timestamptz) language plpgsql security definer set search_path=public as $$
declare original beta_workspaces; next_state jsonb; change jsonb; col text; recid text; old_item jsonb; new_item jsonb; visible_old jsonb; area text; operation text; filtered jsonb; modified integer:=0;
begin
 if not is_member(workspace_org) then raise exception 'Active membership required.' using errcode='42501';end if;
 select * into strict original from beta_workspaces w where w.organization_id=workspace_org and w.id=workspace_id for update;
 if original.updated_at is distinct from expected_updated_at then raise exception 'Workspace changed on another device. Refresh before saving.' using errcode='40001';end if;
 if nl_is_owner(workspace_org) then
 if jsonb_typeof(workspace_payload->'state')<>'object' then raise exception 'Invalid workspace.';end if;
 next_state:=workspace_payload->'state';
 else
 if jsonb_typeof(workspace_payload->'changes') is distinct from 'array' or jsonb_array_length(workspace_payload->'changes')>2000 then raise exception 'Permission-aware changes are required. Reload the admin.' using errcode='42501';end if;
 next_state:=original.payload->'state';
 for change in select value from jsonb_array_elements(workspace_payload->'changes') loop
 col:=change->>'collection';recid:=change->>'id';new_item:=change->'after';
 if recid is null or length(recid)=0 or length(recid)>200 or col is null or nl_collection_area(col)='security' then raise exception 'Invalid record change.' using errcode='42501';end if;
 if exists(select 1 from jsonb_array_elements(workspace_payload->'changes') x where x->>'collection'=col and x->>'id'=recid offset 1) then raise exception 'Duplicate record change.';end if;
 select value into old_item from jsonb_array_elements(coalesce(original.payload->'state'->col,'[]')) where value->>'id'=recid;
 operation:=case when new_item is null or new_item='null'::jsonb then 'delete' when old_item is null then 'create' else 'edit' end;
 if old_item is not null and not coalesce(nl_record_in_scope(workspace_org,original.payload->'state',col,old_item),false) then raise exception 'This record is outside your assigned projects.' using errcode='42501';end if;
 area:=nl_collection_area(col,coalesce(old_item,new_item));
 if not nl_has_permission(workspace_org,area,operation) then raise exception 'Permission denied: % / %.',area,operation using errcode='42501';end if;
 if operation<>'delete' then
 if col='media' then new_item:=new_item||jsonb_build_object('uploadedBy',coalesce(old_item->>'uploadedBy',auth.uid()::text));end if;
 if col='media' and new_item->>'publishStatus'='published' and coalesce(old_item->>'publishStatus','')<>'published' then raise exception 'Website publishing is Owner only.' using errcode='42501';end if;
 if jsonb_typeof(new_item)<>'object' or new_item->>'id' is distinct from recid then raise exception 'Invalid record.';end if;
 if not nl_has_permission(workspace_org,nl_collection_area(col,new_item),operation) or not coalesce(nl_record_in_scope(workspace_org,original.payload->'state',col,new_item),false) then raise exception 'Record group or project is not authorized.' using errcode='42501';end if;
 filtered:=nl_visible_record(workspace_org,col,new_item);
 if filtered<>new_item then raise exception 'Protected fields cannot be changed.' using errcode='42501';end if;
 if old_item is not null then
 visible_old:=nl_visible_record(workspace_org,col,old_item);
 if col='people' and (select scope='assigned' from nl_access_groups g join organization_members m on m.organization_id=g.organization_id and m.role::text=g.group_key where m.organization_id=workspace_org and m.user_id=auth.uid()) and coalesce(old_item->'projectIds','[]') is distinct from coalesce(new_item->'projectIds','[]') then raise exception 'Project assignments are controlled by the office.' using errcode='42501';end if;
 if (coalesce(new_item->>'archived','false')='true' and coalesce(old_item->>'archived','false')<>'true' or new_item->>'status'='void' and coalesce(old_item->>'status','')<>'void') and not nl_has_permission(workspace_org,area,'delete') then raise exception 'Delete permission is required to archive or void this record.' using errcode='42501';end if;
 if col in ('projects','schedule','people') and not nl_has_permission(workspace_org,'financial','edit') and (select coalesce(jsonb_object_agg(key,value),'{}') from jsonb_each(new_item) where key=any(array['contractValue','cost','outstanding','dailyRate','hourlyRate','rate','payRate','balance','totalPaid','units','workValue','dueDate'])) is distinct from (select coalesce(jsonb_object_agg(key,value),'{}') from jsonb_each(visible_old) where key=any(array['contractValue','cost','outstanding','dailyRate','hourlyRate','rate','payRate','balance','totalPaid','units','workValue','dueDate'])) then raise exception 'Financial edit permission is required.' using errcode='42501';end if;
 -- Hidden fields survive changes to visible operational fields.
 new_item:=(old_item-(select coalesce(array_agg(k),'{}') from jsonb_object_keys(visible_old) k))||new_item;
 elsif col in ('projects','schedule','people') and not nl_has_permission(workspace_org,'financial','create') and new_item ?| array['contractValue','cost','outstanding','dailyRate','hourlyRate','rate','payRate','units','workValue'] then raise exception 'Financial create permission is required.' using errcode='42501';end if;
 end if;
 select coalesce(jsonb_agg(value),'[]') into filtered from jsonb_array_elements(coalesce(next_state->col,'[]')) where value->>'id' is distinct from recid;
 if operation<>'delete' then filtered:=filtered||jsonb_build_array(new_item);end if;
 next_state:=jsonb_set(next_state,array[col],filtered,true);modified:=modified+1;
 end loop;
 -- Client project lists are derived from an authorized project update, without granting client edit.
 for change in select value from jsonb_array_elements(workspace_payload->'changes') where value->>'collection'='projects' loop
 recid:=change->>'id';select value into new_item from jsonb_array_elements(coalesce(next_state->'projects','[]')) where value->>'id'=recid;
 select coalesce(jsonb_agg(jsonb_set(c,'{projectIds}',(select coalesce(jsonb_agg(x),'[]') from jsonb_array_elements(coalesce(c->'projectIds','[]')) x where x#>>'{}'<>recid)||case when c->>'id'=new_item->>'clientId' then jsonb_build_array(recid) else '[]'::jsonb end)),'[]') into filtered from jsonb_array_elements(coalesce(next_state->'clients','[]')) c;
 next_state:=jsonb_set(next_state,'{clients}',filtered,true);
 end loop;
 end if;
 update beta_workspaces w set payload=jsonb_build_object('state',next_state,'updatedAt',clock_timestamp(),'updatedBy',auth.uid(),'updatedByRole',(select role from organization_members where organization_id=workspace_org and user_id=auth.uid())),updated_by=auth.uid(),updated_at=clock_timestamp() where w.organization_id=workspace_org and w.id=workspace_id;
 return query select w.updated_at from beta_workspaces w where w.organization_id=workspace_org and w.id=workspace_id;
end; $$;
-- Private objects follow the same department and project rules as their records.
create or replace function public.can_read_workspace_file(bucket text,object_name text) returns boolean language plpgsql stable security definer set search_path=public as $$
declare org uuid; snapshot jsonb; col text; row jsonb;
begin
 if bucket not in ('project-files','compliance-files','media-library') then return false;end if;
 begin org:=split_part(object_name,'/',1)::uuid;exception when invalid_text_representation then return false;end;
 if nl_is_owner(org) then return true;end if;
 if not is_member(org) then return false;end if;
 select nl_project_state(org,w.payload->'state') into snapshot from beta_workspaces w where w.organization_id=org and w.id='shared-v1';
 for col,row in select e.key,r.value from jsonb_each(snapshot) e cross join lateral jsonb_array_elements(e.value) r loop
 if (bucket='media-library' and col='media' and row->>'filePath'=object_name) or (bucket='compliance-files' and col='people' and nl_has_permission(org,'compliance','view') and (row->>'w9FilePath'=object_name or row->>'insuranceFilePath'=object_name)) or (bucket='project-files' and col='expenseReceipts' and row->>'receiptFilePath'=object_name) then return true;end if;
 end loop;
 return false;
end; $$;
create or replace function public.nl_can_upload_file(bucket text,object_name text) returns boolean language plpgsql stable security definer set search_path=public as $$
declare org uuid; area text; folder text; snapshot jsonb; own_person text; scope_value text;
begin
 begin org:=split_part(object_name,'/',1)::uuid;exception when invalid_text_representation then return false;end;
 if nl_is_owner(org) then return bucket in ('project-files','compliance-files','media-library');end if;
 folder:=split_part(object_name,'/',2);
 area:=case bucket when 'media-library' then 'media' when 'compliance-files' then 'compliance' when 'project-files' then 'financial' else '' end;
 if not (nl_has_permission(org,area,'create') or nl_has_permission(org,area,'edit')) then return false;end if;
 select m.linked_person_id,g.scope into own_person,scope_value from organization_members m join nl_access_groups g on g.organization_id=m.organization_id and g.group_key=m.role::text where m.organization_id=org and m.user_id=auth.uid() and m.status='active';
 if scope_value='all' then return true;end if;
 if bucket='compliance-files' then return folder='people' and split_part(object_name,'/',3)=own_person;end if;
 if bucket='project-files' then return folder='receipts' and exists(select 1 from beta_workspaces w,jsonb_array_elements(w.payload->'state'->'people') p where w.organization_id=org and w.id='shared-v1' and p->>'id'=own_person and coalesce(p->'projectIds','[]') ? split_part(object_name,'/',3));end if;
 if bucket='media-library' then return exists(select 1 from beta_workspaces w,jsonb_array_elements(w.payload->'state'->'people') p where w.organization_id=org and w.id='shared-v1' and p->>'id'=own_person and coalesce(p->'projectIds','[]') ? folder);end if;
 return false;
end; $$;

create or replace function public.can_access_project(pid uuid) returns boolean language sql stable security definer set search_path=public as $$ select exists(select 1 from projects p where p.id=pid and nl_is_owner(p.organization_id)); $$;
-- Legacy party/time endpoints are not the operational workspace. Prevent alternate permission bypasses.
drop policy if exists "parties_read" on public.parties;
drop policy if exists "parties_staff" on public.parties;
drop policy if exists "partner_docs_read" on public.partner_documents;
drop policy if exists "partner_docs_staff" on public.partner_documents;
drop policy if exists "assignees_read" on public.schedule_assignees;
drop policy if exists "assignees_staff" on public.schedule_assignees;
drop policy if exists "time_read" on public.time_entries;
drop policy if exists "time_self_insert" on public.time_entries;
drop policy if exists "time_self_update" on public.time_entries;
drop policy if exists "time_staff" on public.time_entries;
create policy nl_owner_legacy on public.parties for all to authenticated using (nl_is_owner(organization_id)) with check (nl_is_owner(organization_id));
create policy nl_owner_legacy on public.partner_documents for all to authenticated using (nl_is_owner(organization_id)) with check (nl_is_owner(organization_id));
create policy nl_owner_legacy on public.schedule_assignees for all to authenticated using (nl_is_owner(organization_id)) with check (nl_is_owner(organization_id));
create policy nl_owner_legacy on public.time_entries for all to authenticated using (nl_is_owner(organization_id)) with check (nl_is_owner(organization_id));

drop policy if exists nl_storage_insert on storage.objects;
create policy nl_storage_insert on storage.objects for insert to authenticated with check (nl_can_upload_file(bucket_id,name));
-- Replacements and file removal require the owning department's explicit Edit/Delete permission.
drop policy if exists nl_storage_update on storage.objects;
create policy nl_storage_update on storage.objects for update to authenticated using (can_read_workspace_file(bucket_id,name) and nl_has_permission(split_part(name,'/',1)::uuid,case bucket_id when 'media-library' then 'media' when 'compliance-files' then 'compliance' else 'financial' end,'edit')) with check (nl_can_upload_file(bucket_id,name));
drop policy if exists nl_storage_delete on storage.objects;
create policy nl_storage_delete on storage.objects for delete to authenticated using (can_read_workspace_file(bucket_id,name) and nl_has_permission(split_part(name,'/',1)::uuid,case bucket_id when 'media-library' then 'media' when 'compliance-files' then 'compliance' else 'financial' end,'delete'));
create or replace function public.nl_guard_owner() returns trigger language plpgsql security definer set search_path=public as $$
begin
 if tg_op='DELETE' then
 if exists(select 1 from nl_access_owner where organization_id=old.organization_id and user_id=old.user_id) then raise exception 'The Owner account is protected.' using errcode='42501';end if;
 return old;
 end if;
 if exists(select 1 from nl_access_owner where organization_id=new.organization_id and user_id=new.user_id) then
 if new.role<>'owner' or new.status<>'active' then raise exception 'The Owner account is protected.' using errcode='42501';end if;
 elsif new.role in ('owner','developer') then raise exception 'Unrestricted access is reserved to the Owner.' using errcode='42501';end if;
 return new;
end; $$;
drop trigger if exists nl_protect_owner on public.organization_members;
create trigger nl_protect_owner before insert or update or delete on public.organization_members for each row execute function public.nl_guard_owner();
CREATE OR REPLACE FUNCTION public.review_visit_request(target_id uuid, actor_id uuid, organization uuid, changes jsonb)
 RETURNS admin_visit_requests
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public'
AS $function$
declare result public.admin_visit_requests; previous public.admin_visit_requests;
begin
  if not public.nl_permission_for(organization,actor_id,'requests',case when changes?'discarded_at' then 'delete' else 'edit' end) or not exists(select 1 from public.organization_members m join public.nl_access_groups g on g.organization_id=m.organization_id and g.group_key=m.role::text where m.organization_id=organization and m.user_id=actor_id and m.status='active' and g.scope='all') and not exists(select 1 from public.nl_access_owner o join public.organization_members m on m.organization_id=o.organization_id and m.user_id=o.user_id where o.organization_id=organization and o.user_id=actor_id and m.role='owner' and m.status='active') then raise exception 'Not authorized' using errcode='42501';end if;
  if exists(select 1 from jsonb_object_keys(changes) k where k not in ('status','internal_note','discarded_at','updated_at')) then raise exception 'Invalid changes'; end if;
  select * into strict previous from public.admin_visit_requests where id=target_id for update;
  update public.admin_visit_requests set
    status=case when changes?'status' then changes->>'status' else status end,
    internal_note=case when changes?'internal_note' then left(changes->>'internal_note',10000) else internal_note end,
    discarded_at=case when changes?'discarded_at' then (changes->>'discarded_at')::timestamptz else discarded_at end,
    updated_at=clock_timestamp()
    where id=target_id returning * into result;
  insert into public.system_audit_events(organization_id,actor_user_id,action,department,entity_type,entity_id,metadata)
  values(organization,actor_id,
    case when changes?'discarded_at' then case when result.discarded_at is null then 'visit_request.restored' else 'visit_request.archived' end else 'visit_request.reviewed' end,
    'administration','visit_request',target_id::text,
    jsonb_build_object('previous_status',previous.status,'status',result.status,'note_changed',previous.internal_note is distinct from result.internal_note));
  return result;
end; $function$;

CREATE OR REPLACE FUNCTION public.get_private_site_report(actor_id uuid, organization uuid, report_period text DEFAULT 'today'::text)
 RETURNS jsonb
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public'
AS $function$
declare report_start timestamptz; report_end timestamptz; local_now timestamp:=clock_timestamp() at time zone 'America/New_York'; cutoff timestamp; result jsonb;
begin
 if organization<>'00000000-0000-4000-8000-000000000001'::uuid or not exists(
  select 1 from public.organization_members where organization_id=organization and user_id=actor_id and status='active' and role='owner' and exists(select 1 from public.nl_access_owner o where o.organization_id=organization and o.user_id=actor_id)
 ) then raise exception 'Your account is not authorized to view this report.' using errcode='42501';end if;
 if report_period='today' then
  report_start:=date_trunc('day',local_now) at time zone 'America/New_York';report_end:=clock_timestamp();
 elsif report_period='yesterday' then
  report_start:=(date_trunc('day',local_now)-interval '1 day') at time zone 'America/New_York';report_end:=date_trunc('day',local_now) at time zone 'America/New_York';
 elsif report_period='week' then
  report_start:=(date_trunc('day',local_now)-interval '6 days') at time zone 'America/New_York';report_end:=clock_timestamp();
 elsif report_period='daily' then
  cutoff:=date_trunc('day',local_now)+interval '20 hours';if local_now<cutoff then cutoff:=cutoff-interval '1 day';end if;
  report_start:=(cutoff-interval '1 day') at time zone 'America/New_York';report_end:=cutoff at time zone 'America/New_York';
 else raise exception 'Invalid report period.' using errcode='22023';end if;
 select report into result from (
with bounds as (select report_start::timestamptz as a,report_end::timestamptz as b),
    metrics as (select m.* from public.site_metric_hourly m,bounds where m.hour>=date_trunc('hour',bounds.a) and m.hour<bounds.b),
    events as (select e.* from public.system_audit_events e,bounds where e.created_at>=bounds.a and e.created_at<bounds.b and e.organization_id='00000000-0000-4000-8000-000000000001' and e.action<>'backup.completed'),
    backups as (select r.* from public.admin_backup_runs r,bounds where r.started_at>=bounds.a and r.started_at<bounds.b and r.organization_id='00000000-0000-4000-8000-000000000001')
    select jsonb_build_object(
      'period_start',report_start,'period_end',report_end,'period',report_period,'timezone','America/New_York',
      'security_controls',jsonb_build_object('mfa_enrolled',exists(select 1 from auth.mfa_factors where user_id=actor_id and status='verified'),'mfa_enforcement','not_verified','file_backup','not_configured_here'),'measurement_started_at',(select started_at from public.site_measurement_config where id=true),
      'traffic',(select coalesce(jsonb_agg(x),'[]'::jsonb) from (select event,dimension,sum(total)::bigint as total from metrics group by event,dimension order by event,total desc)x),
      'hourly_page_views',(select coalesce(jsonb_agg(x),'[]'::jsonb) from (select hour,sum(total)::bigint as total from metrics where event='page_view' group by hour order by hour)x),
      'requests_received',(select count(*) from public.admin_visit_requests r,bounds where r.submitted_at>=bounds.a and r.submitted_at<bounds.b),
      'audit_counts',(select coalesce(jsonb_agg(x),'[]'::jsonb) from (select action,count(*) as total from events group by action order by total desc)x),
      'staff_events',(select coalesce(jsonb_agg(x),'[]'::jsonb) from (select e.created_at,e.action,e.entity_type,coalesce(nullif(p.full_name,''),'Authorized account / system') as actor,e.metadata->'removed_counts' as removed_counts from events e left join public.profiles p on p.id=e.actor_user_id order by e.created_at desc limit 50)x),
      'staff_event_total',(select count(*) from events),
      'backups_completed',(select count(*) from backups where status='completed'),
      'backups_failed',(select count(*) from backups where status='failed'),
      'latest_backup',(select jsonb_build_object('kind',r.backup_kind,'status',r.status,'finished_at',r.finished_at) from public.admin_backup_runs r where r.organization_id='00000000-0000-4000-8000-000000000001' order by r.started_at desc limit 1),
      'limits',jsonb_build_array('Visits are estimated per browser tab; not unique people.','Counts can include automated traffic and omit browsers with privacy signals or blocking.','Hourly buckets provide approximate period boundaries.','Direct/unknown sources include links without a referrer; individual email recipients are not tracked.','Traffic counters cannot determine whether an intrusion occurred. WAF/security-provider logs are not connected.','Staff audit events record account actions, not proof of misconduct.','Backups contain application snapshots and a media manifest; uploaded file bytes require separate recovery.')) as report
 ) report_query;
 return result;
end; $function$;

revoke all on function public.nl_is_owner(uuid) from public,anon;
grant execute on function public.nl_is_owner(uuid) to authenticated;
revoke all on function public.nl_has_permission(uuid,text,text) from public,anon;
grant execute on function public.nl_has_permission(uuid,text,text) to authenticated;
revoke all on function public.nl_access_context(uuid) from public,anon;
grant execute on function public.nl_access_context(uuid) to authenticated;
revoke all on function public.nl_save_access_group(uuid,text,text,jsonb,integer) from public,anon;
grant execute on function public.nl_save_access_group(uuid,text,text,jsonb,integer) to authenticated;
revoke all on function public.nl_set_member_access(uuid,uuid,text,text,text) from public,anon;
grant execute on function public.nl_set_member_access(uuid,uuid,text,text,text) to authenticated;
revoke all on function public.nl_collection_area(text,jsonb) from public,anon;
grant execute on function public.nl_collection_area(text,jsonb) to authenticated;
revoke all on function public.nl_record_in_scope(uuid,jsonb,text,jsonb) from public,anon;
grant execute on function public.nl_record_in_scope(uuid,jsonb,text,jsonb) to authenticated;
revoke all on function public.nl_visible_record(uuid,text,jsonb) from public,anon;
grant execute on function public.nl_visible_record(uuid,text,jsonb) to authenticated;
revoke all on function public.nl_project_state(uuid,jsonb) from public,anon;
grant execute on function public.nl_project_state(uuid,jsonb) to authenticated;
revoke all on function public.nl_can_upload_file(text,text) from public,anon;
grant execute on function public.nl_can_upload_file(text,text) to authenticated;
revoke all on function public.read_operational_workspace(uuid,text) from public,anon;
grant execute on function public.read_operational_workspace(uuid,text) to authenticated;
revoke all on function public.save_admin_workspace(uuid,text,jsonb,timestamptz) from public,anon;
grant execute on function public.save_admin_workspace(uuid,text,jsonb,timestamptz) to authenticated;
notify pgrst, 'reload schema';

grant select on public.nl_access_groups,public.nl_access_owner to service_role;

create or replace function public.nl_intake_permission_for(org uuid,actor uuid,operation text) returns boolean language sql stable security definer set search_path=public as $$
select public.nl_permission_for(org,actor,'requests',operation) and exists(select 1 from organization_members m left join nl_access_groups g on g.organization_id=m.organization_id and g.group_key=m.role::text where m.organization_id=org and m.user_id=actor and m.status='active' and (g.scope='all' or exists(select 1 from nl_access_owner o where o.organization_id=org and o.user_id=actor and m.role='owner'))); $$;
revoke all on function public.nl_intake_permission_for(uuid,uuid,text) from public,anon,authenticated;
grant execute on function public.nl_intake_permission_for(uuid,uuid,text) to service_role;
