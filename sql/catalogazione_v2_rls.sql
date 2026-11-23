-- =====================================================================
-- LogiScan · Catalogazione v2: disattiva la row level security
-- Da eseguire nello SQL Editor di Supabase
--
-- Le tabelle create di recente nascono con RLS attiva e senza policy:
-- la lettura torna vuota e ogni scrittura viene rifiutata con
-- "new row violates row-level security policy".
-- Il resto dello schema del progetto ha la RLS disattivata: queste si allineano.
-- =====================================================================

alter table public.cat_clienti_transcodifica disable row level security;
alter table public.inv_snapshot_magazzini    disable row level security;

-- Create nella v1, ancora presenti: si allineano anche loro
alter table if exists public.cat_sospese disable row level security;

-- Verifica: rowsecurity deve essere false per tutte
select relname as tabella, relrowsecurity as rls_attiva
  from pg_class
 where relname in ('cat_clienti_transcodifica', 'inv_snapshot_magazzini', 'cat_sospese',
                   'cat_impostazioni', 'cat_transcodifiche', 'cat_export_testate',
                   'inv_snapshot', 'inv_snapshot_staging', 'inv_snapshot_riepilogo')
 order by relname;
