-- RightCut — Supabase schema (idempotent migration)
-- Safe to run multiple times — drops and recreates policies

-- Sessions table
create table if not exists public.sessions (
  id          uuid primary key default gen_random_uuid(),
  session_id  text unique not null,
  user_id     uuid references auth.users(id) on delete cascade,
  title       text,
  created_at  timestamptz default now(),
  updated_at  timestamptz default now()
);

-- Messages table
create table if not exists public.messages (
  id          uuid primary key default gen_random_uuid(),
  message_id  text unique not null,
  session_id  text references public.sessions(session_id) on delete cascade,
  user_id     uuid references auth.users(id) on delete cascade,
  role        text not null check (role in ('user','agent','error')),
  text        text,
  metadata    jsonb,
  created_at  timestamptz default now()
);

-- Add metadata column if it doesn't exist (safe for existing tables)
do $$ begin
  alter table public.messages add column if not exists metadata jsonb;
exception when others then null;
end $$;

-- Workbook snapshots table
create table if not exists public.workbook_snapshots (
  id          uuid primary key default gen_random_uuid(),
  session_id  text unique not null references public.sessions(session_id) on delete cascade,
  user_id     uuid references auth.users(id) on delete cascade,
  snapshot    jsonb,
  updated_at  timestamptz default now()
);

-- Add/rename snapshot column if needed (old schema used state_json)
do $$ begin
  alter table public.workbook_snapshots add column if not exists snapshot jsonb;
exception when others then null;
end $$;

-- Enable RLS
alter table public.sessions enable row level security;
alter table public.messages enable row level security;
alter table public.workbook_snapshots enable row level security;

-- Drop all existing policies first (idempotent)
drop policy if exists "Users see own sessions" on public.sessions;
drop policy if exists "Users insert own sessions" on public.sessions;
drop policy if exists "Users update own sessions" on public.sessions;
drop policy if exists "Users delete own sessions" on public.sessions;
drop policy if exists "Users see own messages" on public.messages;
drop policy if exists "Users insert own messages" on public.messages;
drop policy if exists "Users see own snapshots" on public.workbook_snapshots;
drop policy if exists "Users insert own snapshots" on public.workbook_snapshots;
drop policy if exists "Users update own snapshots" on public.workbook_snapshots;

-- Recreate policies
create policy "Users see own sessions"
  on public.sessions for select using (auth.uid() = user_id);
create policy "Users insert own sessions"
  on public.sessions for insert with check (auth.uid() = user_id);
create policy "Users update own sessions"
  on public.sessions for update using (auth.uid() = user_id);
create policy "Users delete own sessions"
  on public.sessions for delete using (auth.uid() = user_id);

create policy "Users see own messages"
  on public.messages for select using (auth.uid() = user_id);
create policy "Users insert own messages"
  on public.messages for insert with check (auth.uid() = user_id);

create policy "Users see own snapshots"
  on public.workbook_snapshots for select using (auth.uid() = user_id);
create policy "Users insert own snapshots"
  on public.workbook_snapshots for insert with check (auth.uid() = user_id);
create policy "Users update own snapshots"
  on public.workbook_snapshots for update using (auth.uid() = user_id);
