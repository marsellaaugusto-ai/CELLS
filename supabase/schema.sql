-- Cells: tablas y permisos. Pegá esto en Supabase → SQL Editor → Run.
-- Se puede volver a ejecutar sin problemas.

create table if not exists public.workbooks (
  id uuid primary key,
  owner uuid not null default auth.uid() references auth.users (id) on delete cascade,
  name text not null,
  data jsonb not null,      -- copia legible del libro (la usan otras apps vía API)
  ydoc text,                -- documento colaborativo (Yjs, base64)
  updated_at timestamptz not null default now()
);
alter table public.workbooks add column if not exists ydoc text;

-- Personas invitadas a editar un libro, por email.
create table if not exists public.workbook_members (
  workbook_id uuid not null references public.workbooks (id) on delete cascade,
  email text not null,
  primary key (workbook_id, email)
);

-- Funciones de permiso (security definer evita que las políticas se llamen entre sí en bucle).
create or replace function public.is_owner(wid uuid) returns boolean
language sql security definer stable set search_path = public as $$
  select exists (select 1 from workbooks where id = wid and owner = auth.uid());
$$;

create or replace function public.can_access(wid uuid) returns boolean
language sql security definer stable set search_path = public as $$
  select exists (
    select 1 from workbooks w
    where w.id = wid
      and (w.owner = auth.uid()
        or exists (select 1 from workbook_members m
                   where m.workbook_id = wid and lower(m.email) = lower(auth.jwt() ->> 'email')))
  );
$$;

alter table public.workbooks enable row level security;
alter table public.workbook_members enable row level security;

drop policy if exists "dueño lee" on public.workbooks;
drop policy if exists "dueño crea" on public.workbooks;
drop policy if exists "dueño edita" on public.workbooks;
drop policy if exists "dueño borra" on public.workbooks;
drop policy if exists "acceso lee" on public.workbooks;
drop policy if exists "acceso edita" on public.workbooks;

create policy "acceso lee" on public.workbooks for select using (public.can_access(id));
create policy "dueño crea" on public.workbooks for insert with check (owner = auth.uid());
create policy "acceso edita" on public.workbooks for update using (public.can_access(id));
create policy "dueño borra" on public.workbooks for delete using (owner = auth.uid());

drop policy if exists "ver invitados" on public.workbook_members;
drop policy if exists "dueño invita" on public.workbook_members;
drop policy if exists "dueño quita" on public.workbook_members;

create policy "ver invitados" on public.workbook_members for select
  using (public.is_owner(workbook_id) or lower(email) = lower(auth.jwt() ->> 'email'));
create policy "dueño invita" on public.workbook_members for insert with check (public.is_owner(workbook_id));
create policy "dueño quita" on public.workbook_members for delete using (public.is_owner(workbook_id));

-- Colaboración en vivo: solo quien tiene acceso al libro entra a su canal "wb:<id>".
drop policy if exists "colaboradores escuchan" on realtime.messages;
drop policy if exists "colaboradores envían" on realtime.messages;

create policy "colaboradores escuchan" on realtime.messages for select to authenticated
  using (realtime.topic() like 'wb:%' and public.can_access(substring(realtime.topic() from 4)::uuid));
create policy "colaboradores envían" on realtime.messages for insert to authenticated
  with check (realtime.topic() like 'wb:%' and public.can_access(substring(realtime.topic() from 4)::uuid));
