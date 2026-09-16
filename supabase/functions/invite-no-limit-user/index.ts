import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": Deno.env.get("ADMIN_APP_ORIGIN") ?? "https://nolimitcontractor.net",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const allowedRoles = new Set(["admin", "manager", "office", "team_member", "vendor", "subcontractor", "viewer"]);

function json(status: number, body: Record<string, unknown>) {
  return new Response(JSON.stringify(body), { headers: { ...corsHeaders, "Content-Type": "application/json" }, status });
}

Deno.serve(async (request) => {
  if (request.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  if (request.method !== "POST") return json(405, { error: "Method not allowed" });

  const supabaseUrl = Deno.env.get("SUPABASE_URL");
  const anonKey = Deno.env.get("SUPABASE_ANON_KEY");
  const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
  const organizationId = Deno.env.get("NO_LIMIT_ORGANIZATION_ID");
  const redirectTo = Deno.env.get("NO_LIMIT_INVITE_REDIRECT_URL");
  const ownerEmail = String(Deno.env.get("NO_LIMIT_OWNER_EMAIL") ?? "").trim().toLowerCase();
  if (!supabaseUrl || !anonKey || !serviceKey || !organizationId || !redirectTo || !ownerEmail) {
    return json(503, { error: "Invitation service is not configured" });
  }

  const authorization = request.headers.get("Authorization") ?? "";
  const callerClient = createClient(supabaseUrl, anonKey, { global: { headers: { Authorization: authorization } }, auth: { persistSession: false } });
  const { data: callerData, error: callerError } = await callerClient.auth.getUser();
  if (callerError || !callerData.user) return json(401, { error: "Sign in again" });

  const admin = createClient(supabaseUrl, serviceKey, { auth: { persistSession: false } });
  const { data: membership } = await admin.from("organization_members").select("role,status")
    .eq("organization_id", organizationId).eq("user_id", callerData.user.id).eq("status", "active").maybeSingle();
  const isVerifiedOwner = String(callerData.user.email ?? "").toLowerCase() === ownerEmail && Boolean(callerData.user.email_confirmed_at);
  if ((!membership || !["owner", "admin"].includes(membership.role)) && !isVerifiedOwner) {
    return json(403, { error: "Only an administrator can manage user access" });
  }

  const body = await request.json().catch(() => ({}));
  const { data: usersPage, error: usersError } = await admin.auth.admin.listUsers({ page: 1, perPage: 1000 });
  if (usersError) return json(500, { error: "Existing accounts could not be checked" });

  if (body.action === "list") {
    const { data: members, error: membersError } = await admin.from("organization_members")
      .select("user_id,role,status").eq("organization_id", organizationId);
    if (membersError) return json(500, { error: "User access records could not be loaded" });
    const usersById = new Map(usersPage.users.map((user) => [user.id, user]));
    return json(200, { users: (members ?? []).map((member) => {
      const user = usersById.get(member.user_id);
      return {
        id: member.user_id,
        email: user?.email ?? "",
        fullName: String(user?.user_metadata?.full_name ?? user?.email ?? ""),
        role: member.role,
        status: member.status,
      };
    }) });
  }

  const email = String(body.email ?? "").trim().toLowerCase();
  const fullName = String(body.fullName ?? "").trim();
  const role = String(body.role ?? "viewer");
  const linkedPersonId = String(body.linkedPersonId ?? "").trim();
  if (!email || !fullName || !allowedRoles.has(role)) return json(400, { error: "Invalid invitation details" });

  let invitedUser = usersPage.users.find((user) => String(user.email ?? "").toLowerCase() === email);
  let invitationSent = false;
  if (!invitedUser) {
    const { data: invited, error: inviteError } = await admin.auth.admin.inviteUserByEmail(email, { redirectTo, data: { full_name: fullName, organization_id: organizationId } });
    if (inviteError || !invited.user) return json(400, { error: inviteError?.message ?? "Invitation failed" });
    invitedUser = invited.user;
    invitationSent = true;
  }

  const { error: profileError } = await admin.from("profiles").upsert({ id: invitedUser.id, email, full_name: fullName }, { onConflict: "id" });
  if (profileError) return json(500, { error: "The account was created, but its profile could not be prepared" });

  const { data: existingMember, error: memberLookupError } = await admin.from("organization_members")
    .select("role,status").eq("organization_id", organizationId).eq("user_id", invitedUser.id).maybeSingle();
  if (memberLookupError) return json(500, { error: "The account exists, but its access could not be checked" });
  if (existingMember?.status === "active") {
    return json(200, { invited: false, accessAssigned: true, alreadyActive: true, role: existingMember.role });
  }

  const { error: memberError } = await admin.from("organization_members").upsert({ organization_id: organizationId, user_id: invitedUser.id, role, status: "active" }, { onConflict: "organization_id,user_id" });
  if (memberError) return json(500, { error: "The account was created, but its role could not be assigned" });

  if (linkedPersonId) {
    const targetTable = role === "vendor" || role === "subcontractor" ? "partners" : "team_members";
    await admin.from(targetTable).update({ user_id: invitedUser.id }).eq("id", linkedPersonId).eq("organization_id", organizationId);
  }
  await admin.from("audit_log").insert({ organization_id: organizationId, actor_user_id: callerData.user.id, action: "user_invited", entity_type: "organization_member", entity_id: invitedUser.id, new_data: { email, full_name: fullName, role, linked_person_id: linkedPersonId || null } });
  return json(200, { invited: invitationSent, accessAssigned: true, alreadyActive: false });
});
