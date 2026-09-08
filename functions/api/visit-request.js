const SUPABASE_URL = 'https://ecbhaeiwnygplawlbnio.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_LDBZuLzFMUo0xsugiKRpZA_qKeqeXeW';

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
    fields = await request.json();
  } catch {
    return json({ error: 'Invalid request.' }, 400);
  }

  const name = String(fields?.name || '').trim();
  const email = String(fields?.email || '').trim();
  if (!name || !email) return json({ error: 'Name and email are required.' }, 400);

  const record = {
    source: 'website',
    page_url: String(fields?.pageUrl || ''),
    name,
    email,
    phone: String(fields?.phone || '').trim() || null,
    address: String(fields?.address || '').trim() || null,
    city: String(fields?.city || '').trim() || null,
    preferred_date: String(fields?.date || '').trim() || null,
    project_type: String(fields?.type || '').trim() || null,
    details: String(fields?.details || '').trim() || null
  };

  try {
    const response = await fetch(`${SUPABASE_URL}/rest/v1/public_visit_requests`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        Prefer: 'return=minimal'
      },
      body: JSON.stringify(record)
    });

    if (!response.ok) {
      console.error('Visit request storage failed.', response.status);
      return json({ error: 'Unable to save request.' }, 502);
    }
  } catch (error) {
    console.error('Visit request service failed.', error);
    return json({ error: 'Unable to save request.' }, 502);
  }

  return json({ ok: true }, 201);
}
