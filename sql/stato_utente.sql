-- Stato corrente di un utente, usato dall'app per rivalidare la sessione salvata
-- in localStorage senza richiedere di nuovo la password.
-- Da eseguire una volta nel SQL Editor di Supabase.
create or replace function public.stato_utente(p_username text)
returns table (username text, ruolo text, attivo boolean, must_change_pw boolean)
language sql
security definer
set search_path = public
as $$
  select u.username::text, u.ruolo::text, u.attivo, u.must_change_pw
    from public.utenti u
   where lower(u.username) = lower(trim(p_username));
$$;

grant execute on function public.stato_utente(text) to anon, authenticated;
