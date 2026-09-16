-- Administrators may list profiles only for people who belong to the same
-- No Limit organization. This supports the private access directory without
-- exposing unrelated Supabase profiles.
drop policy if exists profiles_staff_read on public.profiles;

create policy profiles_staff_read
on public.profiles
for select
to authenticated
using (
  exists (
    select 1
    from public.organization_members target
    where target.user_id = profiles.id
      and public.is_staff(target.organization_id)
  )
);
