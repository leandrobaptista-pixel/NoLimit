CREATE OR REPLACE FUNCTION public.create_admin_backup(target_organization_id uuid, requested_kind text, range_start timestamp with time zone DEFAULT NULL::timestamp with time zone)
 RETURNS uuid
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'storage'
AS $function$
declare run_id uuid; range_end timestamptz:=clock_timestamp(); content jsonb; changed_rows integer:=0;
begin
 if requested_kind not in ('incremental','full') then raise exception 'Unsupported backup kind.' using errcode='22023'; end if;
 insert into public.admin_backup_runs (organization_id,backup_kind,source_from,source_to,created_by) values (target_organization_id,requested_kind,range_start,range_end,auth.uid()) returning id into run_id;
 if requested_kind='incremental' then
  select coalesce(jsonb_agg(jsonb_build_object('workspace_id',version.workspace_id,'source_updated_at',version.source_updated_at,'saved_at',version.saved_at,'change_kind',version.change_kind,'payload',version.payload) order by version.source_updated_at),'[]'::jsonb),count(*) into content,changed_rows from public.admin_workspace_versions version where version.organization_id=target_organization_id and (range_start is null or version.source_updated_at>range_start) and version.source_updated_at<=range_end;
  content:=jsonb_build_object('format','nolimit-admin/incremental-v1','captured_at',range_end,'from',range_start,'to',range_end,'workspace_versions',content);
 else
  select jsonb_build_object('format','nolimit-admin/full-v1','captured_at',range_end,'workspaces',coalesce((select jsonb_agg(jsonb_build_object('workspace_id',workspace.id,'updated_at',workspace.updated_at,'updated_by',workspace.updated_by,'payload',workspace.payload) order by workspace.id) from public.beta_workspaces workspace where workspace.organization_id=target_organization_id),'[]'::jsonb),'members',coalesce((select jsonb_agg(jsonb_build_object('user_id',member.user_id,'role',member.role,'status',member.status) order by member.user_id) from public.organization_members member where member.organization_id=target_organization_id),'[]'::jsonb),'media_manifest',coalesce((select jsonb_agg(jsonb_build_object('bucket_id',storage_object.bucket_id,'name',storage_object.name,'id',storage_object.id,'created_at',storage_object.created_at,'updated_at',storage_object.updated_at,'metadata',storage_object.metadata) order by storage_object.name) from storage.objects storage_object where storage_object.name like target_organization_id::text||'/%'),'[]'::jsonb)) into content;
  select count(*) into changed_rows from public.beta_workspaces workspace where workspace.organization_id=target_organization_id;
 end if;
 content:=content||jsonb_build_object('access_control',jsonb_build_object(
 'owner',(select coalesce(jsonb_agg(to_jsonb(o)),'[]'::jsonb) from public.nl_access_owner o where organization_id=target_organization_id),
 'groups',(select coalesce(jsonb_agg(to_jsonb(g)),'[]'::jsonb) from public.nl_access_groups g where organization_id=target_organization_id),
 'members',(select coalesce(jsonb_agg(to_jsonb(m)),'[]'::jsonb) from public.organization_members m where organization_id=target_organization_id),
 'profiles',(select coalesce(jsonb_agg(to_jsonb(p)),'[]'::jsonb) from public.profiles p join public.organization_members m on m.user_id=p.id where m.organization_id=target_organization_id)));
 update public.admin_backup_runs set status='completed',finished_at=clock_timestamp(),change_count=changed_rows,snapshot=content where id=run_id;
 insert into public.system_audit_events (organization_id,actor_user_id,action,department,entity_type,entity_id,metadata) values (target_organization_id,auth.uid(),'backup.completed','security','backup_run',run_id::text,jsonb_build_object('kind',requested_kind,'change_count',changed_rows,'source_from',range_start,'source_to',range_end));
 return run_id;
end; $function$;
