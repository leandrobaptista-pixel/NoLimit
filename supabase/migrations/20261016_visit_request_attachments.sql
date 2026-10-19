-- Public visitors upload through the validated Edge Function only.
-- No browser SELECT/INSERT policies: files stay private and staff receive expiring links.
insert into storage.buckets (id,name,public,file_size_limit,allowed_mime_types)
values ('visit-request-attachments','visit-request-attachments',false,10485760,
array['image/jpeg','image/png','image/webp','image/heic','application/pdf'])
on conflict (id) do update set public=false,file_size_limit=excluded.file_size_limit,allowed_mime_types=excluded.allowed_mime_types;
