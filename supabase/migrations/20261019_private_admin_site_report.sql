begin;
create or replace function public.get_private_site_report(actor_id uuid,organization uuid,report_period text default 'today')
returns jsonb language plpgsql security definer set search_path=public as $$
declare report_start timestamptz; report_end timestamptz; local_now timestamp:=clock_timestamp() at time zone 'America/New_York'; cutoff timestamp; result jsonb;
begin
 if organization<>'00000000-0000-4000-8000-000000000001'::uuid or not exists(
  select 1 from public.organization_members where organization_id=organization and user_id=actor_id and status='active' and role in ('owner','developer')
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
end; $$;
revoke all on function public.get_private_site_report(uuid,uuid,text) from public,anon,authenticated;
grant execute on function public.get_private_site_report(uuid,uuid,text) to service_role;
commit;
