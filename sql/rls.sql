-- Abilita Row Level Security sulle tabelle del schema public (avviso Supabase
-- "rls_disabled_in_public"). Da eseguire una volta nel SQL Editor di Supabase.
--
-- L'app usa la anon key senza Supabase Auth (login via RPC verifica_login), quindi
-- Postgres non puo' distinguere un utente da un altro: le tabelle che l'app legge
-- e scrive direttamente ricevono una policy aperta per anon/authenticated. Serve a
-- non rompere l'app e a spegnere l'allarme, ma NON protegge i dati da chi ha la
-- anon key. Le tabelle raggiunte solo da RPC security definer (punto 2) invece
-- restano chiuse davvero: RLS attiva e nessuna policy.

-- 1) Tabelle usate direttamente dal client: RLS + policy aperta.
do $$
declare
  t text;
begin
  foreach t in array array[
    'po_lines', 'spare_parts', 'anagrafica_articoli', 'stock_inventory',
    'stock_verifica', 'stock_ordini', 'import_meta', 'refurb_qty',
    'distinte_base', 'distinte_base_tipi',
    'carton_arrivals', 'arrivi_manuali', 'expected_serials', 'scanned_serials',
    'prelievi', 'prelievi_righe',
    'inv_snapshot', 'inv_snapshot_staging', 'inv_snapshot_riepilogo', 'inv_missioni', 'inv_missioni_attese',
    'inv_missioni_rilevazioni', 'inv_missioni_ubicazioni',
    'ml_missioni', 'ml_bancali', 'ml_rilevazioni',
    'cat_impostazioni', 'cat_regole_magazzino', 'cat_transcodifiche',
    'cat_export_testate', 'cat_rma_gestiti', 'cat_sospese', 'cat_clienti_alias'
  ]
  loop
    -- salta nomi inesistenti o che sono viste (alter table fallirebbe)
    if exists (
      select 1 from pg_class c join pg_namespace n on n.oid = c.relnamespace
       where n.nspname = 'public' and c.relname = t and c.relkind in ('r', 'p')
    ) then
      execute format('alter table public.%I enable row level security', t);
      execute format('drop policy if exists app_accesso_aperto on public.%I', t);
      execute format(
        'create policy app_accesso_aperto on public.%I for all to anon, authenticated using (true) with check (true)',
        t);
    else
      raise notice 'tabella public.% non trovata (o non e'' una tabella): saltata', t;
    end if;
  end loop;
end $$;

-- 2) Tabelle accessibili solo tramite RPC security definer: RLS senza policy.
--    Le funzioni girano come owner (postgres) e bypassano RLS, quindi continuano a
--    funzionare; l'accesso diretto con la anon key invece viene negato.
do $$
declare
  t text;
begin
  foreach t in array array['utenti', 'ruoli']
  loop
    if exists (
      select 1 from pg_class c join pg_namespace n on n.oid = c.relnamespace
       where n.nspname = 'public' and c.relname = t and c.relkind in ('r', 'p')
    ) then
      execute format('alter table public.%I enable row level security', t);
    end if;
  end loop;
end $$;

-- 3) Verifica: tabelle di public ancora senza RLS (dovrebbe restituire 0 righe).
--    Se ne compare qualcuna, e' una tabella non coperta sopra: dimmi il nome.
select c.relname as tabella_senza_rls
  from pg_class c join pg_namespace n on n.oid = c.relnamespace
 where n.nspname = 'public' and c.relkind in ('r', 'p') and not c.relrowsecurity
 order by 1;
