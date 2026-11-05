-- =====================================================================
-- LogiScan · Modulo Catalogazione + giacenza NS condivisa
-- Da eseguire una volta nello SQL Editor di Supabase
-- =====================================================================

-- ---------------------------------------------------------------------
-- 1) Giacenza NS: tabella d'appoggio + sostituzione atomica
--    Il CSV viene caricato in inv_snapshot_staging (piena solo durante
--    il caricamento); inv_snapshot_swap() la copia in inv_snapshot e
--    ricostruisce il riepilogo in un'unica transazione, poi la svuota.
--    Chi legge inv_snapshot (missioni, catalogazione) vede sempre la
--    giacenza precedente completa oppure quella nuova completa.
-- ---------------------------------------------------------------------
create table if not exists public.inv_snapshot_staging (like public.inv_snapshot including defaults);

create or replace function public.inv_snapshot_staging_reset()
returns void
language sql
security definer
set search_path = public
as $$
  truncate public.inv_snapshot_staging;
$$;

create or replace function public.inv_snapshot_swap()
returns bigint
language plpgsql
security definer
set search_path = public
set statement_timeout = '300s'
as $$
declare
  n bigint;
begin
  truncate public.inv_snapshot, public.inv_snapshot_riepilogo;
  insert into public.inv_snapshot select * from public.inv_snapshot_staging;
  get diagnostics n = row_count;
  insert into public.inv_snapshot_riepilogo (location, item, qty)
    select location, item, count(*) from public.inv_snapshot group by location, item;
  truncate public.inv_snapshot_staging;
  return n;
end;
$$;

grant all on public.inv_snapshot_staging to anon, authenticated;
grant execute on function public.inv_snapshot_staging_reset() to anon, authenticated;
grant execute on function public.inv_snapshot_swap() to anon, authenticated;

-- ---------------------------------------------------------------------
-- 2) Catalogazione: impostazioni (clienti RIPAX, nomi magazzini)
-- ---------------------------------------------------------------------
create table if not exists public.cat_impostazioni (
  chiave      text primary key,
  valore      jsonb not null,
  updated_at  timestamptz not null default now(),
  updated_by  text
);

insert into public.cat_impostazioni (chiave, valore) values
  ('clienti_ripax', '["19","75","999993","999995"]'::jsonb),
  ('magazzino_ripax', '"RIPAX - PAX DA RIPARARE GESSATE"'::jsonb),
  ('magazzino_rip', '"RIP - TERZI DA RIPARARE GESSATE"'::jsonb)
on conflict (chiave) do nothing;

-- ---------------------------------------------------------------------
-- 3) Catalogazione: regole magazzino origine -> destinazione
--    (assenza di regola = trasferimento normale)
-- ---------------------------------------------------------------------
create table if not exists public.cat_regole_magazzino (
  id                      bigint generated always as identity primary key,
  magazzino_origine       text not null,
  magazzino_destinazione  text not null,
  azione                  text not null check (azione in ('ignora', 'da_verificare')),
  updated_at              timestamptz not null default now(),
  updated_by              text,
  unique (magazzino_origine, magazzino_destinazione)
);

-- ---------------------------------------------------------------------
-- 4) Catalogazione: transcodifica articoli pot_
--    Customer ID + Main Component -> prodotto anagrafica
--    prodotto NULL = "nessun prodotto": righe escluse dagli export
-- ---------------------------------------------------------------------
create table if not exists public.cat_transcodifiche (
  customer_id     text not null,
  main_component  text not null,
  prodotto        text,
  updated_at      timestamptz not null default now(),
  updated_by      text,
  primary key (customer_id, main_component)
);

-- ---------------------------------------------------------------------
-- 5) Catalogazione: storico export (solo testata, nessun dettaglio matricole)
--    idext univoco: impedisce il doppio carico dello stesso documento
-- ---------------------------------------------------------------------
create table if not exists public.cat_export_testate (
  id                      bigint generated always as identity primary key,
  idext                   text not null unique,
  tipo                    text not null check (tipo in ('IT', 'RA')),
  cliente                 text not null,
  data_catalogazione      date not null,
  magazzino_origine       text,
  magazzino_destinazione  text not null,
  n_righe                 integer not null,
  n_matricole             integer not null,
  giacenza_at             timestamptz,
  file_catalogato         text,
  esportato_da            text,
  esportato_at            timestamptz not null default now()
);

grant all on public.cat_impostazioni, public.cat_regole_magazzino, public.cat_transcodifiche, public.cat_export_testate to anon, authenticated;
