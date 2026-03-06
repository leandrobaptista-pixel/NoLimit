# Setup Supabase (Sync Multi-Dispositivo)

Este app usa sync `offline-first`: cada dispositivo grava localmente e sincroniza com Supabase quando online.
O sync cobre `units`, `photos`, `users`, `clients`, `projects`, `contacts` e `containers`.

## 1) Criar tabela

No SQL Editor do Supabase, execute:

```sql
create table if not exists public.app_records (
  tenant text not null,
  kind text not null,
  id text not null,
  payload jsonb not null,
  updated_at timestamptz not null default now(),
  primary key (tenant, kind, id)
);
```

## 2) Politica de acesso (MVP)

Para MVP rapido (sem JWT por usuario), execute:

```sql
alter table public.app_records enable row level security;

create policy "open_select" on public.app_records
for select to anon
using (true);

create policy "open_insert" on public.app_records
for insert to anon
with check (true);

create policy "open_update" on public.app_records
for update to anon
using (true)
with check (true);
```

## 3) Configurar no app

No painel **Sincronizacao na nuvem** do sistema, preencher:

- `Supabase URL`
- `Supabase Anon Key`
- `Tenant/Codigo da empresa` (ex: `tag-nyc`)

Depois:

1. Clique `Enviar para nuvem (Push)` no dispositivo principal.
2. Nos outros dispositivos, use `Baixar da nuvem (Pull)`.
3. Opcional: ative `Auto-sync`.

## 4) Recomendacao de seguranca

Para ambiente definitivo, substituir politica aberta por autenticacao JWT por usuario e regras por `tenant`.
