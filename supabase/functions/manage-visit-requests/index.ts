import { createClient } from "npm:@supabase/supabase-js@2";

const adminUrl = Deno.env.get("SUPABASE_URL")!;
const adminServiceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
// Secret labels are kept short because the Supabase dashboard limits their
// names. Their values stay encrypted in the Admin project's Edge Function
// secret store.
const publicIntakeUrl = Deno.env.get("PUBLIC_INTAKE_SUPA")!;
const publicIntakeKey = Deno.env.get("PUBLIC_INTAKE_SUPABASE_SECRET")!;
const corsHeaders = {
  "Access-Control-Allow-Origin": "https://nolimitcontractor.net",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Content-Type": "application/json",
};

type IntakeRow = {
  id: number; created_at: string; name: string; email: string; phone: string | null;
  address: string | null; city: string | null; preferred_date: string | null;
  project_type: string | null; details: string | null; page_url: string | null;
};

function reply(payload: unknown, status = 200) {
  return new Response(JSON.stringify(payload), { status, headers: corsHeaders });
}

async function requireAuthorizedUser(request: Request) {
  const token = request.headers.get("Authorization")?.replace(/^Bearer\s+/i, "");
  if (!token) throw new Error("Authentication is required.");
  const admin = createClient(adminUrl, adminServiceKey, { auth: { persistSession: false } });
  const { data: { user }, error } = await admin.auth.getUser(token);
  if (error || !user) throw new Error("Authentication has expired. Sign in again.");

  // The directory record is the authoritative role source for the No Limit Admin.
  const { data: records, error: directoryError } = await admin
    .from("app_records")
    .select("payload")
    .eq("tenant", "nolimit")
    .eq("kind", "user");
  if (directoryError) throw directoryError;
  const directoryUser = (records || []).map((record) => record.payload || {}).find((item) =>
    String(item.email || "").toLowerCase() === String(user.email || "").toLowerCase()
  );
  const role = String(directoryUser?.role || "").toLowerCase();
  if (!['owner', 'administrator', 'admin', 'office / financial', 'project manager'].includes(role)) {
    throw new Error("This account is not authorized to review website visit requests.");
  }
  return admin;
}

async function syncWebsiteRequests(admin: ReturnType<typeof createClient>) {
  const response = await fetch(
    `${publicIntakeUrl}/rest/v1/public_visit_requests?select=id,created_at,name,email,phone,address,city,preferred_date,project_type,details,page_url&order=id.asc`,
    { headers: { apikey: publicIntakeKey, Authorization: `Bearer ${publicIntakeKey}` } }
  );
  if (!response.ok) throw new Error("The website intake could not be reached.");
  const rows = await response.json() as IntakeRow[];
  if (!rows.length) return;
  const mapped = rows.map((row) => ({
    source_request_id: row.id, submitted_at: row.created_at, full_name: row.name,
    email: row.email, phone: row.phone, address: row.address, city: row.city,
    preferred_date: row.preferred_date, project_type: row.project_type,
    message: row.details, source_page_url: row.page_url,
  }));
  const { error } = await admin.from("admin_visit_requests").upsert(mapped, { onConflict: "source_request_id" });
  if (error) throw error;
}

Deno.serve(async (request) => {
  if (request.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  try {
    const admin = await requireAuthorizedUser(request);
    const body = await request.json();
    if (body.action === "sync_and_list") {
      await syncWebsiteRequests(admin);
      const { data, error } = await admin.from("admin_visit_requests").select("*").order("submitted_at", { ascending: false });
      if (error) throw error;
      return reply({ requests: data || [] });
    }
    if (body.action === "update") {
      const { data, error } = await admin.from("admin_visit_requests")
        .update({ status: body.status, internal_note: String(body.internalNote || ""), updated_at: new Date().toISOString() })
        .eq("id", body.id).select().single();
      if (error) throw error;
      return reply({ request: data });
    }
    if (body.action === "discard") {
      const { data, error } = await admin.from("admin_visit_requests")
        .update({ discarded_at: body.discarded ? new Date().toISOString() : null, updated_at: new Date().toISOString() })
        .eq("id", body.id).select().single();
      if (error) throw error;
      return reply({ request: data });
    }
    return reply({ error: "Unsupported action." }, 400);
  } catch (error) {
    return reply({ error: error instanceof Error ? error.message : "Request processing failed." }, 400);
  }
});
