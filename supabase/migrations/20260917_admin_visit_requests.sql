-- Private No Limit copy of public website visit requests.
-- Public submissions stay in the website project; only the trusted Edge Function
-- below reads them and synchronizes the minimum operational fields here.

create table if not exists public.admin_visit_requests (
  id uuid primary key default gen_random_uuid(),
  source_request_id bigint not null unique,
  submitted_at timestamptz not null,
  full_name text not null,
  email text not null,
  phone text,
  address text,
  city text,
  preferred_date date,
  project_type text,
  message text,
  source_page_url text,
  status text not null default 'to_contact'
    check (status in ('to_contact', 'waiting_reply', 'in_progress', 'completed')),
  internal_note text not null default '',
  discarded_at timestamptz,
  attachments jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists admin_visit_requests_active_submitted_idx
  on public.admin_visit_requests (discarded_at, submitted_at desc);

create index if not exists admin_visit_requests_project_type_idx
  on public.admin_visit_requests (project_type);

alter table public.admin_visit_requests enable row level security;

-- There are intentionally no browser policies. Access is mediated by the
-- manage-visit-requests Edge Function after it verifies an authenticated user.
