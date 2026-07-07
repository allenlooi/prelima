-- Prelima database schema.
-- Paste this into Supabase: Project -> SQL Editor -> New query -> Run.
-- Safe to re-run any number of times (drops + recreates policies each time).
--
-- Design note: projects and quotes are stored as a single `data jsonb` column
-- (rather than one Postgres column per field) so the app's existing JS object
-- shapes work unchanged. `id`, `user_id` are pulled out as real columns only
-- because they're needed for lookups and Row Level Security.

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  workspace_name text not null default 'My Studio',
  updated_at timestamptz not null default now()
);

create table if not exists public.projects (
  id text primary key,
  user_id uuid not null references auth.users(id) on delete cascade,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists public.quotes (
  id text primary key,
  user_id uuid not null references auth.users(id) on delete cascade,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

create index if not exists projects_user_id_idx on public.projects(user_id);
create index if not exists quotes_user_id_idx on public.quotes(user_id);

alter table public.profiles enable row level security;
alter table public.projects enable row level security;
alter table public.quotes enable row level security;

-- Each freelancer can only ever see/write their own rows.
drop policy if exists "profiles: owner read" on public.profiles;
drop policy if exists "profiles: owner insert" on public.profiles;
drop policy if exists "profiles: owner update" on public.profiles;
create policy "profiles: owner read" on public.profiles for select using (id = auth.uid());
create policy "profiles: owner insert" on public.profiles for insert with check (id = auth.uid());
create policy "profiles: owner update" on public.profiles for update using (id = auth.uid());

drop policy if exists "projects: owner read" on public.projects;
drop policy if exists "projects: owner insert" on public.projects;
drop policy if exists "projects: owner update" on public.projects;
drop policy if exists "projects: owner delete" on public.projects;
create policy "projects: owner read" on public.projects for select using (user_id = auth.uid());
create policy "projects: owner insert" on public.projects for insert with check (user_id = auth.uid());
create policy "projects: owner update" on public.projects for update using (user_id = auth.uid());
create policy "projects: owner delete" on public.projects for delete using (user_id = auth.uid());

drop policy if exists "quotes: owner read" on public.quotes;
drop policy if exists "quotes: owner insert" on public.quotes;
drop policy if exists "quotes: owner update" on public.quotes;
drop policy if exists "quotes: owner delete" on public.quotes;
create policy "quotes: owner read" on public.quotes for select using (user_id = auth.uid());
create policy "quotes: owner insert" on public.quotes for insert with check (user_id = auth.uid());
create policy "quotes: owner update" on public.quotes for update using (user_id = auth.uid());
create policy "quotes: owner delete" on public.quotes for delete using (user_id = auth.uid());
