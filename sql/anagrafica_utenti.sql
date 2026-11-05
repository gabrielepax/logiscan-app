-- Nome e cognome come dati facoltativi dell'utente: servono solo alla schermata
-- admin per riconoscere chi c'è dietro uno username. Nei record e negli export
-- continua a essere salvato esclusivamente lo username.
-- Da eseguire una volta nel SQL Editor di Supabase.

alter table public.utenti
  add column if not exists nome    text,
  add column if not exists cognome text;

-- Elenco esteso: nuova funzione, la elenco_utenti esistente resta invariata.
create or replace function public.elenco_utenti_ext()
returns table (
  username       text,
  ruolo          text,
  nome           text,
  cognome        text,
  attivo         boolean,
  must_change_pw boolean,
  created_at     timestamptz
)
language sql
security definer
set search_path = public
as $$
  select u.username::text, u.ruolo::text, u.nome::text, u.cognome::text,
         u.attivo, u.must_change_pw, u.created_at::timestamptz
    from public.utenti u
   order by u.username;
$$;

grant execute on function public.elenco_utenti_ext() to anon, authenticated;

-- Aggiorna nome/cognome senza toccare password e ruolo. Stringa vuota = campo svuotato.
create or replace function public.set_utente_anagrafica(p_username text, p_nome text, p_cognome text)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
begin
  update public.utenti
     set nome    = nullif(btrim(coalesce(p_nome, '')), ''),
         cognome = nullif(btrim(coalesce(p_cognome, '')), '')
   where lower(username) = lower(btrim(p_username));

  if not found then
    raise exception 'Utente "%" inesistente', p_username;
  end if;

  return true;
end;
$$;

grant execute on function public.set_utente_anagrafica(text, text, text) to anon, authenticated;
