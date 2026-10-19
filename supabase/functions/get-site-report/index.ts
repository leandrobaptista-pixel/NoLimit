import { createClient } from 'npm:@supabase/supabase-js@2';
const headers={'Access-Control-Allow-Origin':'https://nolimitcontractor.net','Access-Control-Allow-Headers':'authorization, x-client-info, apikey, content-type','Access-Control-Allow-Methods':'POST, OPTIONS','Content-Type':'application/json','Cache-Control':'no-store'};
function reply(payload:unknown,status=200){return new Response(JSON.stringify(payload),{status,headers});}
Deno.serve(async request=>{
 if(request.method==='OPTIONS')return new Response(null,{status:204,headers});
 if(request.method!=='POST')return reply({error:'Method not allowed.'},405);
 try {
  const organization=Deno.env.get('NO_LIMIT_ORGANIZATION_ID');if(!organization)return reply({error:'Report access is not configured.'},503);
  const token=/^Bearer\s+(.+)$/i.exec(request.headers.get('Authorization')||'')?.[1]?.trim();
  if(!token)return reply({error:'Sign in to view the private report.'},401);
  const admin=createClient(Deno.env.get('SUPABASE_URL')!,Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,{auth:{persistSession:false}});
  const {data:{user},error:authError}=await admin.auth.getUser(token);
  if(authError||!user)return reply({error:'Sign in again before viewing this report.'},401);
  const {data:member,error:memberError}=await admin.from('organization_members').select('role,status').eq('organization_id',organization).eq('user_id',user.id).eq('status','active').maybeSingle();
  if(memberError)throw memberError;
  if(!member||member.status!=='active'||!['owner','developer'].includes(member.role))return reply({error:'Only an authorized Owner or Developer can view this report.'},403);
  let body;
  try{const text=await request.text();if(text.length>256)return reply({error:'Invalid request.'},400);body=JSON.parse(text);}catch{return reply({error:'Invalid request.'},400);}
  const period=body?.period||'today';
  if(!body||Object.keys(body).some(k=>k!=='period')||!['today','yesterday','week','daily'].includes(period))return reply({error:'Invalid report period.'},400);
  const {data,error}=await admin.rpc('get_private_site_report',{actor_id:user.id,organization,report_period:period});
  if(error)throw error;
  if(!data||typeof data!=='object')throw new Error('Invalid report');
  return reply({report:data});
 } catch {return reply({error:'The private report could not be loaded. Try again.'},503);}
});
