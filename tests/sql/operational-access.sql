-- Run with the migration inside a transaction and roll it back.
do $$ declare r record; begin
 select * into r from public.beta_workspaces where id='shared-v1';
 if r is null then raise exception 'Workspace missing'; end if;
end $$;
select set_config('request.jwt.claim.sub','415b3c7c-2e2c-4b74-ae0c-3a5ee4f36fba',true);
set local role authenticated;
do $$ declare result jsonb; begin
 select payload into result from public.read_operational_workspace('00000000-0000-4000-8000-000000000001','shared-v1');
 if jsonb_array_length(result->'state'->'clients') < 1 then raise exception 'Admin read regression'; end if;
end $$;
reset role;
-- Change one membership only inside this rolled-back transaction to exercise the partner boundary.
update public.organization_members set role='vendor',linked_person_id='PE-SECURITY-TEST'
where user_id='664fe48b-cb30-4e2d-90ba-93ffe17eee97';
update public.beta_workspaces set payload=jsonb_build_object('state',jsonb_build_object(
'people',jsonb_build_array(jsonb_build_object('id','PE-SECURITY-TEST','projectIds',jsonb_build_array('PR-ALLOWED')),jsonb_build_object('id','PE-OTHER','email','private@example.invalid')),
'projects',jsonb_build_array(jsonb_build_object('id','PR-ALLOWED','name','Allowed','contractValue',9999),jsonb_build_object('id','PR-OTHER','name','PRIVATE')),
'clients',jsonb_build_array(jsonb_build_object('name','PRIVATE')),
'media',jsonb_build_array(jsonb_build_object('projectId','PR-ALLOWED','filePath','approved.pdf','publishStatus','approved'),jsonb_build_object('projectId','PR-ALLOWED','filePath','internal.pdf','publishStatus','internal')),
'schedule',jsonb_build_array(jsonb_build_object('personId','PE-SECURITY-TEST','projectId','PR-ALLOWED'),jsonb_build_object('personId','PE-OTHER','projectId','PR-ALLOWED'))
)) where id='shared-v1';
select set_config('request.jwt.claim.sub','664fe48b-cb30-4e2d-90ba-93ffe17eee97',true);
set local role authenticated;
do $$ declare result jsonb; n int; denied boolean:=false; begin
 select payload->'state' into result from public.read_operational_workspace('00000000-0000-4000-8000-000000000001','shared-v1');
 if jsonb_array_length(result->'projects') <> 1 or result->'projects'->0->>'id' <> 'PR-ALLOWED' then raise exception 'Project scope failed'; end if;
 if result::text like '%PRIVATE%' or result::text like '%private@example%' or result::text like '%9999%' or result::text like '%internal.pdf%' then raise exception 'Sensitive data leaked'; end if;
 if jsonb_array_length(result->'schedule') <> 1 then raise exception 'Schedule scope failed'; end if;
 select count(*) into n from public.beta_workspaces; if n <> 0 then raise exception 'Partner can read raw workspace'; end if;
 begin
  perform * from public.save_admin_workspace('00000000-0000-4000-8000-000000000001','shared-v1','{}',now());
 exception when others then denied:=true;
 end;
 if not denied then raise exception 'Partner write was accepted'; end if;
end $$;
reset role;
rollback;
select 'Operational access checks passed; changes rolled back' as result;
