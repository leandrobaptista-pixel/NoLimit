import { createClient } from "npm:@supabase/supabase-js@2";
import { allowedStaffRoles, bearerToken, mapWebsiteRequest } from './intake.ts';

const adminUrl = Deno.env.get('SUPABASE_URL')!;
const adminServiceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
const organizationId = Deno.env.get('NO_LIMIT_ORGANIZATION_ID');
const publicIntakeUrl = Deno.env.get('PUBLIC_INTAKE_SUPA');
const publicIntakeKey = Deno.env.get('PUBLIC_INTAKE_SUPABASE_SECRET');
const corsHeaders = {
  'Access-Control-Allow-Origin': 'https://nolimitcontractor.net',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Content-Type': 'application/json'
};
function reply(payload: unknown, status = 200) {
  return new Response(JSON.stringify(payload), {status, headers: corsHeaders});
}
async function syncWebsiteRequests(admin: ReturnType<typeof createClient>) {
  if (!publicIntakeUrl || !publicIntakeKey) throw new Error('Website intake connection is not configured.');
  for (let offset = 0; ; offset += 500) {
    const query = new URLSearchParams({select:'id,updated_at,payload',tenant:'eq.nolimit',kind:'eq.visitRequest',order:'id.asc',offset:String(offset),limit:'500'});
    const response = await fetch(`${publicIntakeUrl.replace(/\/$/,'')}/rest/v1/app_records?${query}`, {
      headers:{apikey:publicIntakeKey}, signal:AbortSignal.timeout(15000)
    });
    if (!response.ok) throw new Error(`Website intake could not be read (HTTP ${response.status}).`);
    const rows = await response.json();
    if (!Array.isArray(rows)) throw new Error('Website intake returned an invalid response.');
    if (!rows.length) break;
    // Only source fields are updated; staff status, notes, attachments and archive state remain intact.
    const {error} = await admin.from('admin_visit_requests').upsert(rows.map(mapWebsiteRequest), {onConflict:'source_record_id'});
    if (error) throw error;
    if (rows.length < 500) break;
  }
}
Deno.serve(async request => {
  if (request.method === 'OPTIONS') return new Response('ok',{headers:corsHeaders});
  if (request.method !== 'POST') return reply({error:'Method not allowed.'},405);
  try {
    if (!organizationId) return reply({error:'The No Limit organization is not configured.'},503);
    const token = bearerToken(request.headers.get('Authorization'));
    if (!token) return reply({error:'Sign in before reviewing requests.'},401);
    const admin = createClient(adminUrl,adminServiceKey,{auth:{persistSession:false}});
    const {data:{user},error:authError} = await admin.auth.getUser(token);
    if (authError || !user) return reply({error:'Authentication has expired. Sign in again.'},401);
    const {data:member,error:memberError} = await admin.from('organization_members').select('role,status')
      .eq('organization_id',organizationId).eq('user_id',user.id).eq('status','active').maybeSingle();
    if (memberError) throw memberError;
    if (!member || !allowedStaffRoles.has(member.role)) return reply({error:'Your account is not authorized to review website requests.'},403);
    const body = await request.json();
    if (body.action === 'sync_and_list') {
      await syncWebsiteRequests(admin);
      const requests = [];
      for (let offset = 0; ; offset += 500) {
        const {data,error} = await admin.from('admin_visit_requests').select('*').order('submitted_at',{ascending:false}).order('id').range(offset,offset+499);
        if (error) throw error;
        requests.push(...(data || []));
        if (!data || data.length < 500) break;
      }
      return reply({requests});
    }
    if (!/^[0-9a-f-]{36}$/i.test(String(body.id || ''))) return reply({error:'A valid request ID is required.'},400);
    let updates;
    if (body.action === 'update') {
      if (!['to_contact','waiting_reply','in_progress','completed'].includes(body.status)) return reply({error:'Invalid request status.'},400);
      updates = {status:body.status,internal_note:String(body.internalNote || '').slice(0,10000),updated_at:new Date().toISOString()};
    } else if (body.action === 'discard' && typeof body.discarded === 'boolean') {
      updates = {discarded_at:body.discarded ? new Date().toISOString() : null,updated_at:new Date().toISOString()};
    } else return reply({error:'Unsupported action.'},400);
    const {data,error} = await admin.from('admin_visit_requests').update(updates).eq('id',body.id).select().single();
    if (error) throw error;
    return reply({request:data});
  } catch (error) {
    console.error('Visit request operation failed', error instanceof Error ? error.message : 'database error');
    return reply({error:error instanceof Error ? error.message : 'Request processing failed. Refresh and try again.'},400);
  }
});
