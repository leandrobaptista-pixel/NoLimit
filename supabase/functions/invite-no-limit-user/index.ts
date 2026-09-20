import { createClient } from "npm:@supabase/supabase-js@2";

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
  if (!supabaseUrl || !anonKey || !serviceKey || !organizationId || !redirectTo) {
    return json(503, { error: "Invitation service is not configured" });
  }

  const authorization = request.headers.get("Authorization") ?? "";
  const callerClient = createClient(supabaseUrl, anonKey, { global: { headers: { Authorization: authorization } }, auth: { persistSession: false } });
  const { data: callerData, error: callerError } = await callerClient.auth.getUser();
  if (callerError || !callerData.user) return json(401, { error: "Sign in again" });

  const admin = createClient(supabaseUrl, serviceKey, { auth: { persistSession: false } });
  const { data: membership } = await admin.from("organization_members").select("role,status")
    .eq("organization_id", organizationId).eq("user_id", callerData.user.id).eq("status", "active").maybeSingle();
  if (!membership || !["owner", "admin"].includes(membership.role)) {
    return json(403, { error: "Only an administrator can manage user access" });
  }

  const body = await request.json().catch(() => ({}));
  const { data: usersPage, error: usersError } = await admin.auth.admin.listUsers({ page: 1, perPage: 1000 });
  if (usersError) return json(500, { error: "Existing accounts could not be checked" });

  if (body.action === "list") {
    const { data: members, error: membersError } = await admin.from("organization_members")
      .select("user_id,role,status,linked_person_id").eq("organization_id", organizationId);
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
        linkedPersonId: member.linked_person_id || "",
      };
    }) });
  }

  const email = String(body.email ?? "").trim().toLowerCase();
  const fullName = String(body.fullName ?? "").trim();
  const role = String(body.role ?? "viewer");
  const linkedPersonId = String(body.linkedPersonId ?? "").trim();
  if (!email || !fullName || !allowedRoles.has(role)) return json(400, { error: "Invalid invitation details" });

  if (linkedPersonId) {
    const {data: workspace,error:workspaceError} = await admin.from("beta_workspaces").select("payload")
      .eq("organization_id",organizationId).eq("id","shared-v1").single();
    const person = workspace?.payload?.state?.people?.find((p: {id:string}) => p.id === linkedPersonId);
    if (workspaceError || !person) return json(400,{error:"Select an existing person or company from this workspace."});
    if ((role === "vendor" && person.type !== "Vendor") || (role === "subcontractor" && person.type !== "Subcontractor"))
      return json(400,{error:"The account role must match the linked person's record type."});
  }
  if (["vendor","subcontractor","team_member"].includes(role) && !linkedPersonId)
    return json(400,{error:"Link this account to a person or company so its assigned projects can be shared."});
  if (body.action === "validate") return json(200,{valid:true,willSendEmail:!usersPage.users.some(user => user.email?.toLowerCase() === email)});

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

  const { error: memberError } = await admin.from("organization_members").upsert({ organization_id: organizationId, user_id: invitedUser.id, role, status: "active", linked_person_id: linkedPersonId || null }, { onConflict: "organization_id,user_id" });
  if (memberError) return json(500, { error: "The account was created, but its role could not be assigned" });

  await admin.from("audit_log").insert({ organization_id: organizationId, actor_user_id: callerData.user.id, action: "user_invited", entity_type: "organization_member", entity_id: invitedUser.id, new_data: { email, full_name: fullName, role, linked_person_id: linkedPersonId || null } });
  return json(200, { invited: invitationSent, accessAssigned: true, alreadyActive: false });
});
