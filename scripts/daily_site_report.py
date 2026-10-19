#!/usr/bin/env python3
"""Read-only daily No Limit report. Outputs aggregate traffic and private staff audit metadata.
No browser credentials, visitor identities, form contents or backup payloads are printed.
"""
import argparse
import json
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

PROJECT = 'powxzhcscqkhtaicxkyt'
LOCAL = ZoneInfo('America/New_York')

def period(daily=False, now=None):
    end = (now or datetime.now(timezone.utc)).astimezone(LOCAL)
    if daily:
        cutoff = end.replace(hour=20, minute=0, second=0, microsecond=0)
        end = cutoff if end >= cutoff else cutoff-timedelta(days=1)
    return (end-timedelta(days=1)).astimezone(timezone.utc), end.astimezone(timezone.utc)

def query_report(start, end):
    # Values originate from datetime, never user-supplied SQL.
    a, b = start.isoformat(), end.isoformat()
    return f"""with bounds as (select '{a}'::timestamptz as a,'{b}'::timestamptz as b),
    metrics as (select m.* from public.site_metric_hourly m,bounds where m.hour>=date_trunc('hour',bounds.a) and m.hour<bounds.b),
    events as (select e.* from public.system_audit_events e,bounds where e.created_at>=bounds.a and e.created_at<bounds.b and e.organization_id='00000000-0000-4000-8000-000000000001' and e.action<>'backup.completed'),
    backups as (select r.* from public.admin_backup_runs r,bounds where r.started_at>=bounds.a and r.started_at<bounds.b and r.organization_id='00000000-0000-4000-8000-000000000001')
    select jsonb_build_object(
      'period_start','{a}','period_end','{b}','timezone','America/New_York',
      'measurement_started_at',(select started_at from public.site_measurement_config where id=true),
      'traffic',(select coalesce(jsonb_agg(x),'[]'::jsonb) from (select event,dimension,sum(total)::bigint as total from metrics group by event,dimension order by event,total desc)x),
      'hourly_page_views',(select coalesce(jsonb_agg(x),'[]'::jsonb) from (select hour,sum(total)::bigint as total from metrics where event='page_view' group by hour order by hour)x),
      'requests_received',(select count(*) from public.admin_visit_requests r,bounds where r.submitted_at>=bounds.a and r.submitted_at<bounds.b),
      'audit_counts',(select coalesce(jsonb_agg(x),'[]'::jsonb) from (select action,count(*) as total from events group by action order by total desc)x),
      'staff_events',(select coalesce(jsonb_agg(x),'[]'::jsonb) from (select e.created_at,e.action,e.entity_type,coalesce(nullif(p.full_name,''),'Authorized account / system') as actor,e.metadata->'removed_counts' as removed_counts from events e left join public.profiles p on p.id=e.actor_user_id order by e.created_at desc limit 50)x),
      'staff_event_total',(select count(*) from events),
      'backups_completed',(select count(*) from backups where status='completed'),
      'backups_failed',(select count(*) from backups where status='failed'),
      'latest_backup',(select jsonb_build_object('kind',r.backup_kind,'status',r.status,'finished_at',r.finished_at) from public.admin_backup_runs r where r.organization_id='00000000-0000-4000-8000-000000000001' order by r.started_at desc limit 1),
      'limits',jsonb_build_array('Visits are estimated per browser tab; not unique people.','Counts can include automated traffic and omit browsers with privacy signals or blocking.','Hourly buckets provide approximate period boundaries.','Direct/unknown sources include links without a referrer; individual email recipients are not tracked.','Traffic counters cannot determine whether an intrusion occurred. WAF/security-provider logs are not connected.','Staff audit events record account actions, not proof of misconduct.','Backups contain application snapshots and a media manifest; uploaded file bytes require separate recovery.')) as report"""

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--daily',action='store_true',help='Last completed 20:00 to 20:00 New Jersey reporting period.')
    args=parser.parse_args()
    token=os.environ.get('SUPABASE_ACCESS_TOKEN')
    if not token:
        token=(Path.home()/'.supabase/access-token').read_text().strip()
    start,end=period(args.daily)
    request=Request(f'https://api.supabase.com/v1/projects/{PROJECT}/database/query',data=json.dumps({'query':query_report(start,end)}).encode(),headers={'Authorization':'Bearer '+token,'Content-Type':'application/json'},method='POST')
    with urlopen(request,timeout=30) as response: rows=json.load(response)
    if not rows or not isinstance(rows[0].get('report'),dict):raise RuntimeError('Daily report unavailable; do not report missing data as zero.')
    print(json.dumps(rows[0]['report'],ensure_ascii=False,indent=2))

if __name__=='__main__':
    main()
