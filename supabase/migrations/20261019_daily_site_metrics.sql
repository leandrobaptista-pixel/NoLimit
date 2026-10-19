begin;
create table if not exists public.site_metric_hourly (
  hour timestamptz not null,
  event text not null,
  dimension text not null,
  total bigint not null default 0 check(total >= 0),
  primary key(hour,event,dimension),
  check ((event,dimension) in (
    ('page_view','home'),
    ('page_view','catalogs'),
    ('page_view','photo'),
    ('visit_start','website'),
    ('source','direct'),
    ('source','internal'),
    ('source','google'),
    ('source','bing'),
    ('source','social'),
    ('source','email'),
    ('source','other'),
    ('gallery_view','trim'),
    ('gallery_view','wainscoting'),
    ('gallery_view','stairs'),
    ('gallery_view','crown-molding-ceiling-coffered-ceiling'),
    ('gallery_view','decks'),
    ('gallery_view','kitchen-vanities'),
    ('gallery_view','fireplaces-bars'),
    ('gallery_view','outside-doors-windows'),
    ('gallery_view','pergola'),
    ('gallery_view','port-portal'),
    ('gallery_view','commercial'),
    ('gallery_view','trash-container'),
    ('gallery_view','wall-paneling'),
    ('section_view','about'),
    ('section_view','services'),
    ('section_view','gallery'),
    ('section_view','catalogs'),
    ('section_view','process'),
    ('section_view','visit'),
    ('contact_click','phone'),
    ('contact_click','email'),
    ('catalog_open','document'),
    ('form_start','visit'),
    ('form_success','visit'),
    ('form_error','visit')
  ))
);
alter table public.site_metric_hourly enable row level security;
revoke all on public.site_metric_hourly from public,anon,authenticated;
grant select,insert,update,delete on public.site_metric_hourly to service_role;
create table if not exists public.site_measurement_config (
  id boolean primary key default true check(id),
  started_at timestamptz not null default clock_timestamp()
);
insert into public.site_measurement_config(id) values(true) on conflict do nothing;
alter table public.site_measurement_config enable row level security;
revoke all on public.site_measurement_config from public,anon,authenticated;
grant select on public.site_measurement_config to service_role;
create or replace function public.increment_site_metric(metric_event text,metric_dimension text)
returns void language plpgsql security definer set search_path=public as $$
begin
  insert into public.site_metric_hourly(hour,event,dimension,total)
  values(date_trunc('hour',clock_timestamp()),metric_event,metric_dimension,1)
  on conflict(hour,event,dimension) do update set total=least(site_metric_hourly.total+1,1000000);
end; $$;
revoke all on function public.increment_site_metric(text,text) from public,anon,authenticated;
grant execute on function public.increment_site_metric(text,text) to service_role;

-- Preserve the existing version backups and add explicit, content-free removal counts.
create or replace function public.audit_workspace_removals()
returns trigger language plpgsql security definer set search_path=public as $$
declare area text; removed_count integer; removed jsonb := '{}'::jsonb;
begin
  for area in select unnest(array['projects','people','clients','media','estimates','contracts','invoices','schedule','users']) loop
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
drop trigger if exists audit_workspace_removals on public.beta_workspaces;
create trigger audit_workspace_removals after update of payload on public.beta_workspaces for each row execute function public.audit_workspace_removals();

-- Staff request edits and the audit record commit together. Caller is the
-- authenticated staff-checking Edge Function, never an anonymous browser.
create or replace function public.review_visit_request(target_id uuid,actor_id uuid,organization uuid,changes jsonb)
returns public.admin_visit_requests language plpgsql security definer set search_path=public as $$
declare result public.admin_visit_requests; previous public.admin_visit_requests;
begin
  if not exists(select 1 from public.organization_members where user_id=actor_id and organization_id=organization and status='active' and role in ('owner','admin','manager','office')) then
    raise exception 'Not authorized' using errcode='42501';
  end if;
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
end; $$;
revoke all on function public.review_visit_request(uuid,uuid,uuid,jsonb) from public,anon,authenticated;
grant execute on function public.review_visit_request(uuid,uuid,uuid,jsonb) to service_role;
commit;

-- Only old aggregate counters are pruned. Operational data is untouched.
select cron.schedule('nolimit-site-metrics-retention','17 5 * * *',$job$delete from public.site_metric_hourly where hour<now()-interval '90 days';$job$)
where not exists(select 1 from cron.job where jobname='nolimit-site-metrics-retention');
