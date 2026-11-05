-- Soglia di scorta (pezzi) e lotto d'acquisto per combinazione PNIT+SPARE in Distinte Base.
-- La soglia si somma alla stima nel fabbisogno MRP; il lotto e' il multiplo a cui si
-- arrotonda il "Da ordinare" (vuoto = 100, come prima).
-- Da eseguire una volta nel SQL Editor di Supabase.
alter table public.distinte_base
  add column if not exists soglia_scorta integer check (soglia_scorta is null or soglia_scorta >= 0),
  add column if not exists lotto_acquisto integer check (lotto_acquisto is null or lotto_acquisto > 0);
