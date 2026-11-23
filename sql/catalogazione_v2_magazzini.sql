-- =====================================================================
-- LogiScan · Giacenza NS: data di caricamento per magazzino
-- Da eseguire nello SQL Editor di Supabase (dopo catalogazione_v2.sql)
--
-- Con il caricamento additivo la giacenza arriva da più estrazioni
-- (PAX NUOVO, repair, …) con date diverse: qui si tiene traccia, per ogni
-- magazzino, da quale file viene e quando è stato caricato.
-- =====================================================================

create table if not exists public.inv_snapshot_magazzini (
  location    text primary key,
  file_at     timestamptz,   -- data del file NS da cui arriva il magazzino
  caricato_at timestamptz not null default now(),
  caricato_da text
);

grant all on public.inv_snapshot_magazzini to anon, authenticated;

-- Sostituzione additiva: rimpiazza solo i magazzini contenuti nel file e
-- registra la provenienza di ciascuno.
-- La versione senza parametri di catalogazione_v2.sql viene rimossa: resta una sola funzione.
drop function if exists public.inv_snapshot_swap_additivo();

create or replace function public.inv_snapshot_swap_additivo(p_file_at timestamptz default null, p_by text default null)
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

  insert into public.inv_snapshot_magazzini (location, file_at, caricato_at, caricato_da)
    select distinct location, p_file_at, now(), p_by from public.inv_snapshot_staging
  on conflict (location) do update
    set file_at = excluded.file_at, caricato_at = excluded.caricato_at, caricato_da = excluded.caricato_da;

  -- riepilogo ricostruito sull'intera giacenza: un magazzino non presente nel file
  -- può avere perso matricole finite altrove
  truncate public.inv_snapshot_riepilogo;
  insert into public.inv_snapshot_riepilogo (location, item, qty)
    select location, item, count(*) from public.inv_snapshot group by location, item;

  -- magazzini svuotati da questo caricamento: non sono più in giacenza
  delete from public.inv_snapshot_magazzini
   where location not in (select distinct location from public.inv_snapshot);

  truncate public.inv_snapshot_staging;
  return n;
end;
$$;

grant execute on function public.inv_snapshot_swap_additivo(timestamptz, text) to anon, authenticated;
