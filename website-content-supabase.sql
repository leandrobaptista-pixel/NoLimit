create extension if not exists pgcrypto;

create table if not exists website_settings (
  id uuid primary key default gen_random_uuid(),
  company_name text not null,
  phone text not null default '',
  email text not null default '',
  logo_url text not null default '',
  logo_18_years_url text not null default '',
  default_cta text not null default 'Request a Visit',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists website_categories (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  slug text not null unique,
  created_at timestamptz not null default now()
);

create table if not exists website_gallery_items (
  id uuid primary key default gen_random_uuid(),
  title text not null,
  category_id uuid not null references website_categories(id) on delete cascade,
  image_url text not null,
  created_at timestamptz not null default now(),
  published boolean not null default false,
  used_in_social boolean not null default false
);

create index if not exists website_gallery_items_published_created_idx
  on website_gallery_items (published, created_at desc);

alter table website_settings enable row level security;
alter table website_categories enable row level security;
alter table website_gallery_items enable row level security;

drop policy if exists "Public can read website settings" on website_settings;
create policy "Public can read website settings"
  on website_settings
  for select
  using (true);

drop policy if exists "Public can read website categories" on website_categories;
create policy "Public can read website categories"
  on website_categories
  for select
  using (true);

drop policy if exists "Public can read published gallery items" on website_gallery_items;
create policy "Public can read published gallery items"
  on website_gallery_items
  for select
  using (published = true);

comment on table website_settings is 'Main website company/profile settings for the public homepage.';
comment on table website_categories is 'Public gallery categories used by the main website.';
comment on table website_gallery_items is 'Published gallery items rendered on the main website.';

-- Example seed row:
-- insert into website_settings (company_name, phone, email, logo_url, logo_18_years_url, default_cta)
-- values ('No Limit Carpentry', '+1 555 123 4567', 'hello@example.com', 'https://example.com/logo.png', 'https://example.com/logo-18-years.png', 'Request a Visit');
