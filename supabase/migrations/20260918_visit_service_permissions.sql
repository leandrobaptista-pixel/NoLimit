-- Requests remain private: only the authenticated staff-checking server function can use them.
grant select, insert, update on public.admin_visit_requests to service_role;
-- The function checks current, active organization membership before accessing intake.
grant select on public.organization_members to service_role;
