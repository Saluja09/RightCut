-- RightCut — Supabase schema
-- Run this in the Supabase SQL editor: https://supabase.com/dashboard/project/_/sql

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
  created_at  timestamptz default now()
);

-- Workbook snapshots table (optional — for restoring spreadsheet state)
create table if not exists public.workbook_snapshots (
  id          uuid primary key default gen_random_uuid(),
  session_id  text references public.sessions(session_id) on delete cascade,
  user_id     uuid references auth.users(id) on delete cascade,
  state_json  jsonb,
  created_at  timestamptz default now()
);

-- RLS: each user can only see their own data
alter table public.sessions enable row level security;
alter table public.messages enable row level security;
alter table public.workbook_snapshots enable row level security;

-- Sessions policies
create policy "Users see own sessions"
  on public.sessions for select using (auth.uid() = user_id);
create policy "Users insert own sessions"
  on public.sessions for insert with check (auth.uid() = user_id);
create policy "Users update own sessions"
  on public.sessions for update using (auth.uid() = user_id);

-- Messages policies
create policy "Users see own messages"
  on public.messages for select using (auth.uid() = user_id);
create policy "Users insert own messages"
  on public.messages for insert with check (auth.uid() = user_id);

-- Snapshots policies
create policy "Users see own snapshots"
  on public.workbook_snapshots for select using (auth.uid() = user_id);
create policy "Users insert own snapshots"
  on public.workbook_snapshots for insert with check (auth.uid() = user_id);

-- Enable anonymous sign-ins (required for "Continue as guest")
-- Go to: Authentication > Providers > Anonymous > Enable
