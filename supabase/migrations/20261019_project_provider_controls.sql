begin;
-- Provider rates and cash ledgers never enter partner payloads.
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
 'schedule',coalesce((select jsonb_agg(jsonb_build_object(
 'id',p->>'id','personId',p->>'personId','personName',p->>'personName','personType',p->>'personType',
 'projectId',p->>'projectId','projectName',p->>'projectName','workDate',p->>'workDate',
 'shift',p->>'shift','instructions',p->>'instructions','status',p->>'status')) from jsonb_array_elements(coalesce(snapshot->'schedule','[]'::jsonb)) p
 where p->>'personId'=person_id and p->>'projectId' in(select id from ids)),'[]'::jsonb),
 'media',coalesce((select jsonb_agg(p) from jsonb_array_elements(coalesce(snapshot->'media','[]'::jsonb)) p
 where p->>'projectId' in(select id from ids) and p->>'publishStatus' in ('approved','published')),'[]'::jsonb)
);
$$;
revoke all on function public.partner_workspace_projection(jsonb,text) from public;

create or replace function public.audit_workspace_removals()
returns trigger language plpgsql security definer set search_path=public as $$
declare area text; removed_count integer; removed jsonb := '{}'::jsonb;
begin
  for area in select unnest(array['projects','people','clients','media','estimates','contracts','invoices','schedule','users','projectEngagements','providerMilestones','providerPayments','payments']) loop
    select count(*) into removed_count
    from jsonb_array_elements(coalesce(old.payload->'state'->area,'[]'::jsonb)) old_item
    where old_item->>'id' is not null and not exists (
      select 1 from jsonb_array_elements(coalesce(new.payload->'state'->area,'[]'::jsonb)) new_item
      where new_item->>'id'=old_item->>'id'
    );
    if removed_count>0 then removed:=removed||jsonb_build_object(area,removed_count); end if;
  end loop;
  if removed<>'{}'::jsonb then
    insert into public.system_audit_events(organization_id,actor_user_id,action,department,entity_type,entity_id,metadata)
    values(new.organization_id,coalesce(new.updated_by,auth.uid()),'workspace.records_removed','security','workspace',new.id,
      jsonb_build_object('removed_counts',removed,'version_at',new.updated_at));
  end if;
  return new;
end; $$;

-- Preserve new collections when an older open browser saves unrelated fields.
-- Recorded provider payments are append-only, with a documented void transition.
create or replace function public.guard_project_cash_ledgers()
returns trigger language plpgsql set search_path=public as $$
declare area text; previous jsonb; replacement jsonb; record_id text;
begin
 for area in select unnest(array['projectEngagements','providerMilestones','providerPayments']) loop
  if not (coalesce(new.payload->'state','{}'::jsonb) ? area) then
   new.payload:=jsonb_set(new.payload,array['state',area],coalesce(old.payload->'state'->area,'[]'::jsonb),true);
  end if;
  if jsonb_typeof(new.payload->'state'->area) <> 'array' then
   raise exception 'Project financial collections must be arrays';
  end if;
 end loop;
 for previous in select value from jsonb_array_elements(coalesce(old.payload->'state'->'providerPayments','[]'::jsonb)) loop
  record_id:=previous->>'id';
  select value into replacement from jsonb_array_elements(new.payload->'state'->'providerPayments') where value->>'id'=record_id;
  if replacement is null then raise exception 'Provider payments cannot be removed; void with a reason'; end if;
  if replacement<>previous then
   if previous->>'status'<>'paid' or replacement->>'status'<>'void'
     or length(trim(coalesce(replacement->>'voidReason','')))=0
     or (replacement - array['status','voidReason','voidedAt','voidedBy'])<>(previous - array['status','voidReason','voidedAt','voidedBy']) then
    raise exception 'Provider payments are immutable; void with a reason and create a corrected entry';
   end if;
   replacement:=replacement||jsonb_build_object('voidedAt',now(),'voidedBy',coalesce(new.updated_by,auth.uid()));
   select jsonb_agg(case when value->>'id'=record_id then replacement else value end) into previous
   from jsonb_array_elements(new.payload->'state'->'providerPayments');
   new.payload:=jsonb_set(new.payload,'{state,providerPayments}',previous);
  end if;
 end loop;
 return new;
end; $$;
revoke all on function public.guard_project_cash_ledgers() from public;
drop trigger if exists guard_project_cash_ledgers on public.beta_workspaces;
create trigger guard_project_cash_ledgers before update of payload on public.beta_workspaces for each row execute function public.guard_project_cash_ledgers();
commit;
