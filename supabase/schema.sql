-- Cells: tabla de libros. Pegá esto en Supabase → SQL Editor → Run.

create table if not exists public.workbooks (
  id uuid primary key,
  owner uuid not null default auth.uid() references auth.users (id) on delete cascade,
  name text not null,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.workbooks enable row level security;

-- Cada usuario ve y edita solo sus libros.
create policy "dueño lee" on public.workbooks for select using (owner = auth.uid());
create policy "dueño crea" on public.workbooks for insert with check (owner = auth.uid());
create policy "dueño edita" on public.workbooks for update using (owner = auth.uid());
create policy "dueño borra" on public.workbooks for delete using (owner = auth.uid());

-- Avisos en tiempo real cuando otro dispositivo o app guarda un libro.
alter publication supabase_realtime add table public.workbooks;
