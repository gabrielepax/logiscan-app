-- =====================================================================
-- LogiScan · Missioni Libere: rilevazione matricolare senza lista attesa
-- Da eseguire una volta nello SQL Editor di Supabase
--
-- Il magazziniere apre una missione, spara l'etichetta del bancale
-- (PLT-GGMMAA-NNN) e poi le matricole (singole o QR master del cartone).
-- Al download dell'Excel resta la testata con i conteggi; il dettaglio
-- (bancali e matricole) può essere eliminato dal database.
-- Stati: in_corso → completata → chiusa (dettaglio eliminato)
-- =====================================================================

create table if not exists public.ml_missioni (
  id              bigserial primary key,
  descrizione     text,                 -- riferimento libero (cliente, DDT, motivo...)
  stato           text not null default 'in_corso',
  creata_da       text,
  created_at      timestamptz not null default now(),
  completata_at   timestamptz,
  -- esiti di testata: calcolati alla chiusura dal dettaglio, restano dopo la pulizia
  n_matricole     integer,
  n_bancali       integer,
  n_segnalazioni  integer,              -- matricole già presenti in altre missioni libere
  scaricata_da    text,
  scaricata_at    timestamptz
);

create table if not exists public.ml_bancali (
  missione_id  bigint not null references public.ml_missioni(id) on delete cascade,
  bancale      text not null,
  operatore    text,
  sparato_at   timestamptz not null default now(),
  primary key (missione_id, bancale)
);

create table if not exists public.ml_rilevazioni (
  id           bigserial primary key,
  missione_id  bigint not null references public.ml_missioni(id) on delete cascade,
  matricola    text not null,
  bancale      text not null,
  modello      text,                    -- dal QR master del cartone, se presente
  fonte        text,                    -- 'singola' | 'qr'
  nota         text,
  operatore    text,
  rilevata_at  timestamptz not null default now(),
  unique (missione_id, matricola)
);

create index if not exists ml_rilevazioni_matricola_idx on public.ml_rilevazioni (matricola);

grant all on public.ml_missioni, public.ml_bancali, public.ml_rilevazioni to anon, authenticated;
grant usage, select on sequence public.ml_missioni_id_seq, public.ml_rilevazioni_id_seq to anon, authenticated;
