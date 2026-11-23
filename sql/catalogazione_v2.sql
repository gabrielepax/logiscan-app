-- =====================================================================
-- LogiScan · Catalogazione v2: controllo backlog repair contro giacenza NS
-- Da eseguire nello SQL Editor di Supabase
-- =====================================================================

-- ---------------------------------------------------------------------
-- 1) Caricamento giacenza ADDITIVO per magazzino
--    Le estrazioni NS sono più di una e coprono magazzini diversi
--    (PAX NUOVO da una parte, RIP/RIPAX/DEP*/IRREP dall'altra):
--    un caricamento sostituisce SOLO i magazzini contenuti nel file,
--    gli altri restano. Vengono tolte anche le matricole che nel file
--    risultano spostate da un magazzino non presente nel file stesso,
--    altrimenti resterebbero duplicate.
-- ---------------------------------------------------------------------
create or replace function public.inv_snapshot_swap_additivo()
returns bigint
language plpgsql
security definer
set search_path = public
set statement_timeout = '300s'
as $$
declare
  n bigint;
begin
  delete from public.inv_snapshot
   where location in (select distinct location from public.inv_snapshot_staging)
      or numero   in (select numero from public.inv_snapshot_staging);

  insert into public.inv_snapshot select * from public.inv_snapshot_staging;
  get diagnostics n = row_count;

  -- riepilogo ricostruito sull'intera giacenza: un magazzino non presente nel file
  -- può avere perso matricole finite altrove
  truncate public.inv_snapshot_riepilogo;
  insert into public.inv_snapshot_riepilogo (location, item, qty)
    select location, item, count(*) from public.inv_snapshot group by location, item;

  truncate public.inv_snapshot_staging;
  return n;
end;
$$;

grant execute on function public.inv_snapshot_swap_additivo() to anon, authenticated;

-- ---------------------------------------------------------------------
-- 2) Transcodifica cliente: nome nel backlog -> Customer ID Anagrafica
--    Il backlog esporta il nome ("FABRICK S.P.A.", "PAX TOREF"), l'Anagrafica
--    ragiona per Customer ID. Il Customer ID determina anche la destinazione:
--    RIPAX se è tra i clienti RIPAX delle impostazioni, altrimenti RIP.
-- ---------------------------------------------------------------------
create table if not exists public.cat_clienti_transcodifica (
  nome_backlog  text primary key,
  customer_id   text not null,
  note          text,
  updated_at    timestamptz not null default now(),
  updated_by    text
);

grant all on public.cat_clienti_transcodifica to anon, authenticated;

-- ---------------------------------------------------------------------
-- 3) Tabelle della versione precedente, non più usate dal modulo.
--    Eseguire SOLO quando si è sicuri di non volerne più i dati.
-- ---------------------------------------------------------------------
-- drop table if exists public.cat_rma_gestiti;
-- drop function if exists public.cat_rma_registra(jsonb, text);
-- drop table if exists public.cat_sospese;
-- drop table if exists public.cat_regole_magazzino;
-- drop table if exists public.cat_clienti_alias;
