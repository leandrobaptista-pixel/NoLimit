begin;

-- These records deliberately never enter beta_workspaces, shared reports or backups.
create table public.nl_credit_members (
  person text primary key check (person in ('leandro','geanniny')),
  user_id uuid not null unique references auth.users(id)
);
create table public.nl_credit_settings (
  id boolean primary key default true check(id),
  budget numeric(14,2) not null default 0 check(budget>=0 and budget<>'NaN'::numeric),
  revision integer not null default 1
);
insert into public.nl_credit_settings(id) values(true);
create table public.nl_credit_debts (
  id uuid primary key default gen_random_uuid(),
  name text not null check(length(trim(name)) between 1 and 120),
  kind text not null check(kind in ('card','mortgage','auto','loan','other')),
  in_plan boolean not null default false check(not in_plan or kind='card'),
  opening_balance numeric(14,2) not null check(opening_balance>=0 and opening_balance<>'NaN'::numeric),
  apr numeric(9,6) not null check(apr between 0 and 10),
  minimum numeric(14,2) check(minimum>=0 and minimum<>'NaN'::numeric),
  due_date date,
  promo_balance numeric(14,2) check(promo_balance>=0 and promo_balance<>'NaN'::numeric),
  promo_deadline date,
  promo_installments integer check(promo_installments>0),
  notes text not null default '' check(length(notes)<=2000),
  archived boolean not null default false,
  revision integer not null default 1,
  created_by uuid not null default auth.uid() references auth.users(id),
  updated_at timestamptz not null default now()
);
create unique index nl_credit_debt_name on public.nl_credit_debts(lower(trim(name)));
create table public.nl_credit_entries (
  id uuid primary key default gen_random_uuid(),
  request_id uuid not null unique default gen_random_uuid(),
  debt_id uuid not null references public.nl_credit_debts(id),
  posted_on date not null,
  payment numeric(14,2) not null default 0 check(payment>=0 and payment<>'NaN'::numeric),
  interest numeric(14,2) not null default 0 check(interest>=0 and interest<>'NaN'::numeric),
  charges numeric(14,2) not null default 0 check(charges>=0 and charges<>'NaN'::numeric),
  invoice_due date,
  notes text not null default '' check(length(notes)<=2000),
  voided_at timestamptz,
  created_by uuid not null default auth.uid() references auth.users(id),
  created_at timestamptz not null default now(),
  check(payment+interest+charges>0)
);
create index nl_credit_entries_debt_date on public.nl_credit_entries(debt_id,posted_on);
create table public.nl_credit_audit (
  id bigint generated always as identity primary key,
  actor_id uuid not null references auth.users(id),
  action text not null,
  record_id uuid,
  created_at timestamptz not null default now()
);
alter table public.nl_credit_members enable row level security;
alter table public.nl_credit_settings enable row level security;
alter table public.nl_credit_debts enable row level security;
alter table public.nl_credit_entries enable row level security;
alter table public.nl_credit_audit enable row level security;
-- No direct Data API access, including for other Owners or Developers.
revoke all on public.nl_credit_members,public.nl_credit_settings,public.nl_credit_debts,public.nl_credit_entries,public.nl_credit_audit from public,anon,authenticated;

create function public.nl_credit_allowed() returns boolean
language sql stable security definer set search_path='' as $$
  select auth.uid() is not null and exists(
    select 1 from public.nl_credit_members a
    join public.organization_members m on m.user_id=a.user_id
    where a.user_id=auth.uid() and m.organization_id='00000000-0000-4000-8000-000000000001'
      and m.status='active'
  );
$$;

create function public.nl_credit_read() returns jsonb
language plpgsql security definer set search_path='' as $$
begin
  if not public.nl_credit_allowed() then raise exception 'Private credit access denied.' using errcode='42501';end if;
  return jsonb_build_object(
    'settings',(select to_jsonb(s) from public.nl_credit_settings s where id),
    'debts',(select coalesce(jsonb_agg(d order by d.name),'[]'::jsonb) from public.nl_credit_debts d),
    'entries',(select coalesce(jsonb_agg(e order by e.posted_on desc,e.created_at desc),'[]'::jsonb) from public.nl_credit_entries e)
  );
end;
$$;

create function public.nl_credit_write(operation text, item jsonb) returns uuid
language plpgsql security definer set search_path='' as $$
declare target uuid; expected integer; affected integer; result_id uuid; request_key uuid; previous public.nl_credit_entries;
begin
  if not public.nl_credit_allowed() then raise exception 'Private credit access denied.' using errcode='42501';end if;
  if jsonb_typeof(item) is distinct from 'object' then raise exception 'An entry is required.';end if;
  -- Serialize edits/imports. Revisions protect debt details and budget from stale writes.
  perform pg_advisory_xact_lock(21112026,1);
  target:=nullif(item->>'id','')::uuid;
  expected:=nullif(item->>'revision','')::integer;
  if operation='budget' then
    if not(item?'budget') or item->>'budget' is null then raise exception 'Enter a monthly budget.';end if;
    update public.nl_credit_settings set budget=(item->>'budget')::numeric,revision=revision+1 where id and revision=expected;
    get diagnostics affected=row_count;
    if affected<>1 then raise exception 'Budget changed. Refresh and try again.' using errcode='40001';end if;
  elsif operation='debt' then
    if not(item?'opening_balance' and item?'apr' and item?'name' and item?'kind') or
       item->>'opening_balance' is null or item->>'apr' is null then raise exception 'Complete the required debt fields.';end if;
    if target is null then
      insert into public.nl_credit_debts(name,kind,in_plan,opening_balance,apr,minimum,due_date,promo_balance,promo_deadline,promo_installments,notes)
      values(trim(item->>'name'),item->>'kind',coalesce((item->>'in_plan')::boolean,false),(item->>'opening_balance')::numeric,(item->>'apr')::numeric,nullif(item->>'minimum','')::numeric,nullif(item->>'due_date','')::date,nullif(item->>'promo_balance','')::numeric,nullif(item->>'promo_deadline','')::date,nullif(item->>'promo_installments','')::integer,coalesce(item->>'notes','')) returning id into result_id;
    else
      update public.nl_credit_debts set name=trim(item->>'name'),kind=item->>'kind',in_plan=coalesce((item->>'in_plan')::boolean,false),opening_balance=(item->>'opening_balance')::numeric,apr=(item->>'apr')::numeric,
        minimum=nullif(item->>'minimum','')::numeric,due_date=nullif(item->>'due_date','')::date,promo_balance=nullif(item->>'promo_balance','')::numeric,
        promo_deadline=nullif(item->>'promo_deadline','')::date,promo_installments=nullif(item->>'promo_installments','')::integer,
        notes=coalesce(item->>'notes',''),revision=revision+1,updated_at=now()
      where id=target and revision=expected returning id into result_id;
      if result_id is null then raise exception 'Debt changed. Refresh and try again.' using errcode='40001';end if;
    end if;
  elsif operation='archive' then
    update public.nl_credit_debts set archived=not archived,revision=revision+1,updated_at=now()
      where id=target and revision=expected returning id into result_id;
    if result_id is null then raise exception 'Debt changed. Refresh and try again.' using errcode='40001';end if;
  elsif operation='entry' then
    request_key:=coalesce(nullif(item->>'request_id','')::uuid,gen_random_uuid());
    select * into previous from public.nl_credit_entries where request_id=request_key;
    if found then
      if previous.debt_id is distinct from (item->>'debt_id')::uuid or previous.posted_on is distinct from (item->>'posted_on')::date
        or previous.payment is distinct from coalesce(nullif(item->>'payment','')::numeric,0)
        or previous.interest is distinct from coalesce(nullif(item->>'interest','')::numeric,0)
        or previous.charges is distinct from coalesce(nullif(item->>'charges','')::numeric,0)
        or previous.invoice_due is distinct from nullif(item->>'invoice_due','')::date
        or previous.notes is distinct from coalesce(item->>'notes','') then raise exception 'This entry was already posted. Refresh before making a correction.';end if;
      return previous.id;
    end if;
    if not exists(select 1 from public.nl_credit_debts where id=(item->>'debt_id')::uuid and not archived) then raise exception 'Choose an active debt.';end if;
    if (item->>'posted_on')::date > (now() at time zone 'America/New_York')::date then raise exception 'Only post completed transactions.';end if;
    insert into public.nl_credit_entries(request_id,debt_id,posted_on,payment,interest,charges,invoice_due,notes)
    values(request_key,(item->>'debt_id')::uuid,(item->>'posted_on')::date,coalesce(nullif(item->>'payment','')::numeric,0),coalesce(nullif(item->>'interest','')::numeric,0),coalesce(nullif(item->>'charges','')::numeric,0),nullif(item->>'invoice_due','')::date,coalesce(item->>'notes','')) returning id into result_id;
  elsif operation='void' then
    update public.nl_credit_entries set voided_at=now() where id=target and voided_at is null returning id into result_id;
    if result_id is null then raise exception 'Entry is unavailable or already voided.';end if;
  else raise exception 'Unknown operation.';
  end if;
  insert into public.nl_credit_audit(actor_id,action,record_id) values(auth.uid(),operation,result_id);
  return result_id;
end;
$$;

-- First import is atomic and only allowed into an empty ledger. A retry cannot duplicate data.
create function public.nl_credit_import(document jsonb) returns integer
language plpgsql security definer set search_path='' as $$
declare debt jsonb; entry jsonb; debt_uuid uuid; total integer:=0;
begin
  if not public.nl_credit_allowed() then raise exception 'Private credit access denied.' using errcode='42501';end if;
  perform pg_advisory_xact_lock(21112026,1);
  if exists(select 1 from public.nl_credit_debts) or exists(select 1 from public.nl_credit_entries) then raise exception 'Initial import is only allowed into an empty credit module.';end if;
  if document->>'format' is distinct from 'nolimit-credit-v1' or jsonb_typeof(document->'debts') is distinct from 'array'
     or jsonb_typeof(document->'entries') is distinct from 'array' then raise exception 'Invalid import file.';end if;
  if jsonb_array_length(document->'debts')>500 or jsonb_array_length(document->'entries')>10000 then raise exception 'Import exceeds supported size.';end if;
  perform public.nl_credit_write('budget',jsonb_build_object('budget',document->'budget','revision',(select revision from public.nl_credit_settings where id)));
  for debt in select value from jsonb_array_elements(document->'debts') loop
    if debt?'id' then raise exception 'Imported debts cannot supply an ID.';end if;
    perform public.nl_credit_write('debt',debt);total:=total+1;
  end loop;
  for entry in select value from jsonb_array_elements(document->'entries') loop
    select id into strict debt_uuid from public.nl_credit_debts where name=entry->>'debt_name';
    perform public.nl_credit_write('entry',entry||jsonb_build_object('debt_id',debt_uuid));
  end loop;
  return total;
end;
$$;
revoke all on function public.nl_credit_allowed(),public.nl_credit_read(),public.nl_credit_write(text,jsonb),public.nl_credit_import(jsonb) from public,anon,authenticated;
grant execute on function public.nl_credit_allowed(),public.nl_credit_read(),public.nl_credit_write(text,jsonb),public.nl_credit_import(jsonb) to authenticated;
commit;
