-- =====================================================================
-- LogiScan · Catalogazione v2: trasferimenti (Inventory Transfer)
-- Da eseguire nello SQL Editor di Supabase
--
-- Le matricole del backlog che NetSuite ha già, ma su un magazzino
-- diverso dalla destinazione attesa per quel cliente, non si caricano
-- con una RA: si spostano con un IT. Per ogni coppia di magazzini
-- l'utente decide una volta cosa fare, e la scelta resta.
--
--   segnala     = compare nei Trasferimenti ma fuori dall'export (default)
--   trasferisci = entra nel file IT
--   ignora      = non se ne parla più
-- =====================================================================

create table if not exists public.cat_regole_magazzini (
  origine      text not null,   -- codice breve del magazzino NS di origine, es. WOFF
  destinazione text not null,   -- codice breve del magazzino di destinazione, es. RIPAX
  azione       text not null default 'segnala'
               check (azione in ('segnala', 'trasferisci', 'ignora')),
  updated_at   timestamptz not null default now(),
  updated_by   text,
  primary key (origine, destinazione)
);

alter table public.cat_regole_magazzini disable row level security;
grant all on public.cat_regole_magazzini to anon, authenticated;

comment on table public.cat_regole_magazzini is
  'Catalogazione: azione per coppia magazzino di origine -> destinazione nei trasferimenti (IT)';

-- Lo storico delle testate accoglie anche i documenti IT (tipo = 'IT',
-- magazzino_origine valorizzato). Un IT è un movimento tra magazzini e non ha
-- un cliente, mentre la colonna era nata obbligatoria con le sole RA.
alter table public.cat_export_testate alter column cliente drop not null;
