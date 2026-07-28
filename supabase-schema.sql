-- À coller dans Supabase → SQL Editor → Run

create table app_storage (
  key text primary key,
  value jsonb not null,
  shared boolean not null default true,
  updated_at timestamptz not null default now()
);

-- Accès ouvert via la clé publique (anon) — acceptable pour un usage interne
-- à une petite équipe. On pourra durcir ça plus tard si besoin (mots de passe
-- réels, accès restreint par équipe, etc.).
alter table app_storage enable row level security;
create policy "allow all" on app_storage for all using (true) with check (true);
