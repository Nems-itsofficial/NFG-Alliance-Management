-- Alliance Command — Supabase schema
-- Run this once in Supabase: Dashboard > SQL Editor > New query > paste all > Run.

create extension if not exists pgcrypto;

-- Single-row settings table (alliance name, leader, thresholds)
create table if not exists settings (
  id int primary key default 1,
  alliance_name text default '',
  leader_name text default '',
  inactivity_days int default 10,
  leaver_retention_days int default 90,
  constraint settings_singleton check (id = 1)
);
insert into settings (id) values (1) on conflict (id) do nothing;

create table if not exists members (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  game_id text default '',
  rank text default 'R1',
  status text not null default 'active' check (status in ('active', 'inactive', 'left')),
  join_date date,
  left_date date,
  notes text default '',
  created_at timestamptz default now()
);

-- One row per member, upserted in place (mirrors the app's "player profile" model)
create table if not exists growth (
  member_id uuid primary key references members(id) on delete cascade,
  power numeric,
  previous_power numeric,
  furnace_level text default '',
  classes jsonb default '{}'::jsonb,
  updated_date date
);

create table if not exists events (
  id uuid primary key default gen_random_uuid(),
  date date not null,
  type text not null,
  name text not null,
  session text default ''
);

create table if not exists participation (
  id uuid primary key default gen_random_uuid(),
  event_id uuid references events(id) on delete cascade,
  member_id uuid references members(id) on delete cascade,
  signed_up boolean default false,
  attended boolean default false,
  partial boolean default false,
  score numeric,
  note text default '',
  unique (event_id, member_id)
);

-- ---------------------------------------------------------------
-- Row Level Security: only signed-in officers (created manually in
-- Authentication > Users, see README) can read or write anything.
-- There is no public/anonymous access.
-- ---------------------------------------------------------------
alter table settings enable row level security;
alter table members enable row level security;
alter table growth enable row level security;
alter table events enable row level security;
alter table participation enable row level security;

create policy "officers full access" on settings for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "officers full access" on members for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "officers full access" on growth for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "officers full access" on events for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "officers full access" on participation for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
