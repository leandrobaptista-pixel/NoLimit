create table if not exists public.public_visit_requests (
  id bigint generated always as identity primary key,
  created_at timestamptz not null default now(),
  source text not null default 'website',
  page_url text,
  name text not null,
  email text not null,
  phone text,
  address text,
  city text,
  preferred_date date,
  project_type text,
  details text
);

alter table public.public_visit_requests enable row level security;

create policy "public_visit_requests_insert"
on public.public_visit_requests
for insert
to anon
with check (true);
