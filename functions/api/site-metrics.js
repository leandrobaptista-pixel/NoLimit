const endpoint = 'https://powxzhcscqkhtaicxkyt.supabase.co/functions/v1/record-site-metrics';
const key = 'sb_publishable_Gw7E39Yd7wR3C0llQavA2w_909B4VpG';
export const dimensions = {
  page_view:['home','catalogs','photo'], visit_start:['website'], source:['direct','internal','google','bing','social','email','other'],
  gallery_view:['trim','wainscoting','stairs','crown-molding-ceiling-coffered-ceiling','decks','kitchen-vanities','fireplaces-bars','outside-doors-windows','pergola','port-portal','commercial','trash-container','wall-paneling'],
  section_view:['about','services','gallery','catalogs','process','visit'], contact_click:['phone','email'],catalog_open:['document'],form_start:['visit'],form_success:['visit'],form_error:['visit']
};
export async function onRequestPost({request}) {
  if (!['https://nolimitcontractor.net','https://www.nolimitcontractor.net'].includes(request.headers.get('origin'))) return new Response(null,{status:403});
  if (request.headers.get('sec-gpc') === '1' || request.headers.get('dnt') === '1') return new Response(null,{status:204});
  if (!(request.headers.get('content-type') || '').startsWith('application/json')) return new Response(null,{status:415});
  // Bound streamed bodies as well as clients that supply Content-Length.
  let bytes=0; const chunks=[];
  const reader=request.body?.getReader(); if(!reader)return new Response(null,{status:400});
  while(true){const {value,done}=await reader.read();if(done)break;bytes+=value.byteLength;if(bytes>256){await reader.cancel();return new Response(null,{status:413});}chunks.push(value);}
  let body;
  try { const data=new Uint8Array(bytes);let offset=0;for(const chunk of chunks){data.set(chunk,offset);offset+=chunk.byteLength;}body=JSON.parse(new TextDecoder().decode(data)); }
  catch{return new Response(null,{status:400});}
  if (!body || Object.keys(body).some(k=>!['event','dimension'].includes(k)) || !dimensions[body.event]?.includes(body.dimension)) return new Response(null,{status:400});
  try {
    const response=await fetch(endpoint,{method:'POST',headers:{'Content-Type':'application/json',apikey:key,Authorization:`Bearer ${key}`},body:JSON.stringify({event:body.event,dimension:body.dimension})});
    return new Response(null,{status:response.ok?204:503,headers:{'Cache-Control':'no-store'}});
  } catch { return new Response(null,{status:503}); }
}
export function onRequestGet(){return new Response(null,{status:405});}
