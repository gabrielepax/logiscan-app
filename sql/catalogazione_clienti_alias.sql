-- =====================================================================
-- LogiScan · Catalogazione: corrispondenza cliente catalogato -> cliente Anagrafica
-- Da eseguire una volta nello SQL Editor di Supabase
--
-- Alcuni clienti del catalogato hanno i prodotti intestati al capogruppo
-- (es. 1355 -> 44, 1705 -> 281, 19 -> 29). Con la corrispondenza salvata
-- il controllo sul Customer ID non segnala più l'errore e la ricerca dei
-- prodotti pot_ usa il cliente giusto.
-- =====================================================================

create table if not exists public.cat_clienti_alias (
  cliente             text primary key,   -- Customer del catalogato (es. 1355)
  cliente_anagrafica  text not null,      -- Customer ID in Anagrafica (es. 44)
  updated_at          timestamptz not null default now(),
  updated_by          text
);

grant all on public.cat_clienti_alias to anon, authenticated;
