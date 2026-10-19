begin;
do $$
declare result jsonb;
begin
 result:=public.partner_workspace_projection('{"projects":[{"id":"p","name":"House","contractValue":10000}],"people":[{"id":"v","projectIds":["p"]},{"id":"other"}],"schedule":[{"id":"s","personId":"v","projectId":"p","workDate":"2026-10-19","dailyRate":999,"units":1,"workValue":999,"dueDate":"2026-10-20"},{"id":"hidden","personId":"other","projectId":"p"}],"projectEngagements":[{"id":"secret"}],"providerPayments":[{"amount":999}],"providerMilestones":[{"amount":999}]}','v');
 if result::text like '%999%' or result::text like '%secret%' or result::text like '%hidden%' or result::text like '%contractValue%' then raise exception 'Partner financial leak'; end if;
 if jsonb_array_length(result->'schedule')<>1 or result->'schedule'->0->>'id'<>'s' then raise exception 'Partner schedule missing'; end if;
 if jsonb_array_length(public.partner_workspace_projection('{"people":[],"projects":[],"schedule":[]}',null)->'projects')<>0 then raise exception 'Unlinked partner access'; end if;
end $$;
create temporary table cash_guard_test(payload jsonb,updated_by uuid);
create trigger cash_guard_test before update on cash_guard_test for each row execute function public.guard_project_cash_ledgers();
insert into cash_guard_test values ('{"state":{"projects":[],"projectEngagements":[{"id":"e"}],"providerPayments":[{"id":"pay","engagementId":"e","amount":300,"status":"paid"}],"providerMilestones":[]}}',null);
-- Simulate an old browser that does not yet know the new collections.
update cash_guard_test set payload='{"state":{"projects":[]}}';
do $$ begin
 if (select payload->'state'->'providerPayments'->0->>'amount' from cash_guard_test)<>'300' then raise exception 'Legacy browser lost cash ledger'; end if;
end $$;
do $$
declare rejected boolean:=false;
begin
 begin update cash_guard_test set payload=jsonb_set(payload,'{state,providerPayments,0,amount}','999'); exception when others then rejected:=true; end;
 if not rejected then raise exception 'Cash amount mutation accepted'; end if;
 rejected:=false;
 begin update cash_guard_test set payload=jsonb_set(payload,'{state,providerPayments}','[]'); exception when others then rejected:=true; end;
 if not rejected then raise exception 'Cash deletion accepted'; end if;
end $$;
update cash_guard_test set payload=jsonb_set(payload,'{state,providerPayments,0}',(payload->'state'->'providerPayments'->0)||'{"status":"void","voidReason":"Test correction","voidedAt":"2026-10-19","voidedBy":"test"}'::jsonb);
do $$ begin
 if (select payload->'state'->'providerPayments'->0->>'status' from cash_guard_test)<>'void' then raise exception 'Documented void failed'; end if;
end $$;
rollback;
