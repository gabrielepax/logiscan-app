-- =====================================================================
-- LogiScan · Catalogazione v2: anagrafica clienti NetSuite
-- Da eseguire nello SQL Editor di Supabase
--
-- Serve per mostrare il nome accanto al Customer ID (l'Internal ID da solo
-- è illeggibile) e per proporre da sé l'abbinamento con i nomi del backlog.
-- Si alimenta con l'export NetSuite dei Customers (Internal ID, ID, Name, [PAX] GROUP).
-- =====================================================================

create table if not exists public.cat_clienti_ns (
  internal_id  text primary key,   -- è il valore usato in [PAX] Customer ID e nel campo Customer della RA
  codice       text,               -- colonna "ID" dell'export: codice secondario, non usarlo per le associazioni
  nome         text not null,
  gruppo       text,
  updated_at   timestamptz not null default now(),
  updated_by   text
);

alter table public.cat_clienti_ns disable row level security;
grant all on public.cat_clienti_ns to anon, authenticated;
