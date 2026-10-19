import { createClient } from 'npm:@supabase/supabase-js@2';
const dimensions = {
  page_view:['home','catalogs','photo'], visit_start:['website'], source:['direct','internal','google','bing','social','email','other'],
  gallery_view:['trim','wainscoting','stairs','crown-molding-ceiling-coffered-ceiling','decks','kitchen-vanities','fireplaces-bars','outside-doors-windows','pergola','port-portal','commercial','trash-container','wall-paneling'],
  section_view:['about','services','gallery','catalogs','process','visit'], contact_click:['phone','email'],catalog_open:['document'],form_start:['visit'],form_success:['visit'],form_error:['visit']
};

Deno.serve(async request => {
  if(request.method!=='POST')return new Response(null,{status:405});
  if(Number(request.headers.get('content-length'))>256)return new Response(null,{status:413});
  let body;
  try {
    let bytes=0;const chunks=[];const reader=request.body?.getReader();
    if(!reader)return new Response(null,{status:400});
    while(true){const {value,done}=await reader.read();if(done)break;bytes+=value.byteLength;if(bytes>256){await reader.cancel();return new Response(null,{status:413});}chunks.push(value);}
    const data=new Uint8Array(bytes);let offset=0;for(const chunk of chunks){data.set(chunk,offset);offset+=chunk.byteLength;}
    body=JSON.parse(new TextDecoder().decode(data));
  } catch{return new Response(null,{status:400});}
  if(!body || Object.keys(body).some(k=>!['event','dimension'].includes(k)) || !dimensions[body.event as keyof typeof dimensions]?.includes(body.dimension))return new Response(null,{status:400});
  const admin=createClient(Deno.env.get('SUPABASE_URL')!,Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,{auth:{persistSession:false}});
  const {error}=await admin.rpc('increment_site_metric',{metric_event:body.event,metric_dimension:body.dimension});
  return new Response(null,{status:error?503:204});
});
