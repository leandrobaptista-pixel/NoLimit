-- Existing server-side invitation service only. User requested full operational release
-- after the prior explicit explanation of these persistent service permissions.
grant select, insert, update on public.profiles to service_role;
grant select, insert, update on public.organization_members to service_role;
grant insert on public.audit_log to service_role;
