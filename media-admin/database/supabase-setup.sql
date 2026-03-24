create extension if not exists pgcrypto;

-- Run schema first.
-- This companion script configures Supabase Storage for the media admin module.

insert into storage.buckets (id, name, public)
values ('nolimit-media', 'nolimit-media', true)
on conflict (id) do update set public = excluded.public;

drop policy if exists "Public can read No Limit media objects" on storage.objects;
create policy "Public can read No Limit media objects"
  on storage.objects
  for select
  to public
  using (bucket_id = 'nolimit-media');

-- Service-role uploads do not require extra object policies.
-- If you later want browser-direct uploads, add authenticated insert/update policies here.
