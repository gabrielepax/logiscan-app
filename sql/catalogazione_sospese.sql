-- =====================================================================
-- LogiScan · Catalogazione: righe messe da parte perché in errore
-- Da eseguire una volta nello SQL Editor di Supabase
--
-- All'export le righe con errori non bloccano più il resto: vengono
-- salvate qui con il motivo. Rientrano da sole nelle analisi successive
-- (anche senza ricaricare il catalogato) e, quando l'errore è risolto,
-- vengono esportate come tutte le altre e tolte da questa tabella.
-- =====================================================================

create table if not exists public.cat_sospese (
  rma                 text not null,
  serial              text not null,
  customer            text,
  data_catalogazione  text,           -- gg/mm/aaaa come nel catalogato
  your_ref            text,
  item                text,
  esito               text,           -- esito calcolato quando è stata messa da parte
  errori              text,           -- motivi, separati da " · "
  file_catalogato     text,
  messa_at            timestamptz not null default now(),
  messa_da            text,
  primary key (rma, serial)
);

grant all on public.cat_sospese to anon, authenticated;
