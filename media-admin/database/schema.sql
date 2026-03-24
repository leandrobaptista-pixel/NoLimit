create extension if not exists pgcrypto;

create table if not exists categories (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  slug text not null unique,
  created_at timestamptz not null default now()
);

create table if not exists media_items (
  id uuid primary key default gen_random_uuid(),
  title text not null,
  category_id uuid not null references categories(id) on delete restrict,
  original_filename text not null,
  original_path text not null,
  original_url text not null,
  generated_art_path text,
  generated_art_url text,
  caption_text text,
  commercial_copy text,
  mime_type text not null,
  file_size_bytes bigint not null default 0,
  published boolean not null default false,
  used_in_social boolean not null default false,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists media_items_category_created_idx
  on media_items (category_id, created_at desc);

create index if not exists media_items_published_idx
  on media_items (published, created_at desc);
