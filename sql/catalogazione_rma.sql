-- =====================================================================
-- LogiScan · Catalogazione: storico RMA + seriali già caricati su NetSuite
-- Da eseguire una volta nello SQL Editor di Supabase
--
-- Un RMA per riga con l'elenco dei seriali: 714.400 combinazioni stanno
-- in ~18.000 righe (~12 MB) invece di ~80 MB con una riga per combinazione.
-- I seriali si conservano come sono (zeri iniziali compresi); il confronto
-- ignora gli zeri iniziali, così un catalogato che li ha persi combacia lo stesso.
-- =====================================================================

create table if not exists public.cat_rma_gestiti (
  rma            text primary key,
  serials        text[] not null default '{}',
  n_serials      integer not null default 0,
  aggiornato_at  timestamptz not null default now(),
  aggiornato_by  text
);

-- Registra/fonde in un'unica transazione le coppie RMA + seriali.
-- p: [{"rma": "000511_2026-09-01_1", "serials": ["0822360405", "..."]}, ...]
-- Ritorna il numero di seriali effettivamente aggiunti (quelli già presenti non contano).
create or replace function public.cat_rma_registra(p jsonb, p_by text default null)
returns integer
language plpgsql
security definer
set search_path = public
as $$
declare
  voce      jsonb;
  v_rma     text;
  v_serials text[];
  aggiunti  integer := 0;
  prima     integer;
begin
  for voce in select * from jsonb_array_elements(p) loop
    v_rma := upper(trim(voce ->> 'rma'));
    select array(select distinct upper(trim(s)) from jsonb_array_elements_text(voce -> 'serials') s
                 where trim(s) <> '') into v_serials;
    if v_rma is null or v_rma = '' or coalesce(array_length(v_serials, 1), 0) = 0 then
      continue;
    end if;
    select coalesce(n_serials, 0) into prima from cat_rma_gestiti where rma = v_rma;

    insert into cat_rma_gestiti as t (rma, serials, n_serials, aggiornato_at, aggiornato_by)
    values (v_rma, v_serials, array_length(v_serials, 1), now(), p_by)
    on conflict (rma) do update
      set serials = (select array(select distinct e from unnest(t.serials || excluded.serials) e)),
          aggiornato_at = now(),
          aggiornato_by = p_by;

    update cat_rma_gestiti set n_serials = coalesce(array_length(serials, 1), 0) where rma = v_rma;
    select n_serials - coalesce(prima, 0) into prima from cat_rma_gestiti where rma = v_rma;
    aggiunti := aggiunti + greatest(prima, 0);
  end loop;
  return aggiunti;
end;
$$;

grant all on public.cat_rma_gestiti to anon, authenticated;
grant execute on function public.cat_rma_registra(jsonb, text) to anon, authenticated;
