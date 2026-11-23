-- =====================================================================
-- LogiScan · Catalogazione v2: carico diretto del Main Component
-- Da eseguire nello SQL Editor di Supabase
--
-- Per alcuni clienti (tipicamente quelli interni come PAX DOA e PAX TOREF,
-- Internal ID 217 = PAX ITALIA) non serve il prodotto/assembly intestato
-- al cliente: si carica direttamente il Main Component del backlog, che in
-- Anagrafica esiste come articolo del cluster Hardware.
-- =====================================================================

alter table public.cat_clienti_transcodifica
  add column if not exists usa_main_component boolean not null default false;

comment on column public.cat_clienti_transcodifica.usa_main_component is
  'true = nella RA si usa il Main Component del backlog come Item, senza cercare il prodotto del cliente';
