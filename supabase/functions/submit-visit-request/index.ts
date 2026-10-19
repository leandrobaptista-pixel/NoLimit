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
  let admin: ReturnType<typeof createClient> | undefined;
  const uploaded: string[] = [];
  try {
    if (Number(request.headers.get('content-length')) > 26 * 1048576) return reply({error:'Files exceed the 25 MB total limit.'},413);
    const multipart = (request.headers.get('content-type') || '').includes('multipart/form-data');
    const form = multipart ? await request.formData() : null;
    const fields = form ? Object.fromEntries(form.entries()) : await request.json();
    const files = form ? form.getAll('attachments').filter((value): value is File => typeof value !== 'string' && value.size > 0) : [];
    if (files.length > 5 || files.some(file => file.size > 10 * 1048576) || files.reduce((sum, file) => sum + file.size, 0) > 25 * 1048576) {
      return reply({error:'Attach up to 5 files, 10 MB each and 25 MB total.'},413);
    }
    const validated = [];
    for (const file of files) {
      const bytes = new Uint8Array(await file.slice(0,32).arrayBuffer());
      const text = new TextDecoder().decode(bytes);
      const ext = file.name.split('.').pop()?.toLowerCase();
      const mime = ext === 'pdf' && text.startsWith('%PDF-') ? 'application/pdf'
        : ['jpg','jpeg'].includes(ext || '') && bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255 ? 'image/jpeg'
        : ext === 'png' && [137,80,78,71,13,10,26,10].every((v,i) => bytes[i] === v) ? 'image/png'
        : ext === 'webp' && text.startsWith('RIFF') && text.slice(8,12) === 'WEBP' ? 'image/webp'
        : ['heic','heif'].includes(ext || '') && text.slice(4,8) === 'ftyp' && /heic|heix|hevc|hevx|mif1|msf1/.test(text.slice(8)) ? 'image/heic' : '';
      if (!mime) return reply({error:'One file is not a supported JPG, PNG, WebP, HEIC or PDF. Please check the file and try again.'},415);
      validated.push({file,mime,ext});
    }
    const name = String(fields?.name || '').trim();
    const email = String(fields?.email || '').trim();
    if (!name || name.length > 200 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) {
      return reply({error:'A valid name and email are required.'},400);
    }
    const id = crypto.randomUUID();
    const date = String(fields?.date || '');
    admin = createClient(Deno.env.get('SUPABASE_URL')!,Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,{auth:{persistSession:false}});
    const attachments = [];
    for (const {file,mime,ext} of validated) {
      const path = `${id}/${crypto.randomUUID()}.${ext}`;
      const {error} = await admin.storage.from('visit-request-attachments').upload(path, file, {contentType:mime,upsert:false});
      if (error) throw error;
      uploaded.push(path);
      attachments.push({bucket:'visit-request-attachments',path,name:file.name.replace(/[\x00-\x1f\x7f]/g,'').slice(0,200),type:mime,size:file.size});
    }
    const {error} = await admin.from('admin_visit_requests').insert({
      attachments,
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
    if (admin && uploaded.length) await admin.storage.from('visit-request-attachments').remove(uploaded).catch(() => {});
    return reply({error:'Unable to save request. Please try again.'},502);
  }
});
