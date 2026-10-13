import { createClient } from 'npm:@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': 'https://nolimitcontractor.net',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Content-Type': 'application/json'
};
function reply(body: unknown, status: number) {
  return new Response(JSON.stringify(body), {status, headers: corsHeaders});
}
Deno.serve(async request => {
  if (request.method === 'OPTIONS') return new Response(null, {status:204, headers:corsHeaders});
  if (request.method !== 'POST') return reply({error:'Method not allowed.'},405);
  try {
    const fields = await request.json();
    const name = String(fields?.name || '').trim();
    const email = String(fields?.email || '').trim();
    if (!name || name.length > 200 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) {
      return reply({error:'A valid name and email are required.'},400);
    }
    const id = crypto.randomUUID();
    const date = String(fields?.date || '');
    const admin = createClient(Deno.env.get('SUPABASE_URL')!,Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,{auth:{persistSession:false}});
    const {error} = await admin.from('admin_visit_requests').insert({
      source_record_id:id, submitted_at:new Date().toISOString(), full_name:name, email,
      phone:String(fields?.phone || '').trim().slice(0,100),
      address:String(fields?.address || '').trim().slice(0,1000),
      city:String(fields?.city || '').trim().slice(0,200),
      preferred_date:/^\d{4}-\d{2}-\d{2}$/.test(date) && Number.isFinite(Date.parse(date)) ? date : null,
      project_type:String(fields?.type || '').trim().slice(0,200),
      message:String(fields?.details || '').trim().slice(0,10000),
      source_page_url:String(fields?.pageUrl || '').slice(0,2000)
    });
    if (error) throw error;
    return reply({ok:true,id},201);
  } catch {
    return reply({error:'Unable to save request. Please try again.'},502);
  }
});
