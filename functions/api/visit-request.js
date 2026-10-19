const SUPABASE_URL = 'https://powxzhcscqkhtaicxkyt.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_Gw7E39Yd7wR3C0llQavA2w_909B4VpG';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type'
};

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      ...corsHeaders,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
}

export async function onRequestOptions() {
  return new Response(null, { status: 204, headers: corsHeaders });
}

export async function onRequestGet() {
  return json({ error: 'Method not allowed.' }, 405);
}

export async function onRequestPost({ request }) {
  let fields;

  try {
    if (Number(request.headers.get('content-length')) > 26 * 1048576) return json({error:'Files exceed the 25 MB total limit.'},413);
    const multipart = (request.headers.get('content-type') || '').includes('multipart/form-data');
    fields = multipart ? await request.formData() : await request.json();
  } catch {
    return json({ error: 'Invalid request.' }, 400);
  }

  const name = String(fields instanceof FormData ? fields.get('name') || '' : fields?.name || '').trim();
  const email = String(fields instanceof FormData ? fields.get('email') || '' : fields?.email || '').trim();
  if (!name || !email) return json({ error: 'Name and email are required.' }, 400);

  try {
    const response = await fetch(`${SUPABASE_URL}/functions/v1/submit-visit-request`, {
      method: 'POST',
      headers: {
        ...(fields instanceof FormData ? {} : {'Content-Type': 'application/json'}),
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
      },
      body: fields instanceof FormData ? fields : JSON.stringify(fields)
    });

    if (!response.ok) {
      console.error('Visit request storage failed.', response.status);
      if ([400,413,415].includes(response.status)) {
        const result = await response.json().catch(() => ({}));
        return json({error: result.error || 'Please check your files and try again.'}, response.status);
      }
      return json({ error: 'Unable to save request.' }, 502);
    }
  } catch (error) {
    console.error('Visit request service failed.', error);
    return json({ error: 'Unable to save request.' }, 502);
  }

  return json({ ok: true }, 201);
}
