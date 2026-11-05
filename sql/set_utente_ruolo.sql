-- Cambia il ruolo di un utente esistente senza doverne reimpostare la password.
-- Da eseguire una volta nel SQL Editor di Supabase.
create or replace function public.set_utente_ruolo(p_username text, p_ruolo text)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare
  v_ruolo text;
begin
  select nome into v_ruolo from public.ruoli where lower(nome) = lower(trim(p_ruolo));
  if v_ruolo is null then
    raise exception 'Ruolo "%" inesistente', p_ruolo;
  end if;

  update public.utenti
     set ruolo = v_ruolo
   where lower(username) = lower(trim(p_username));

  if not found then
    raise exception 'Utente "%" inesistente', p_username;
  end if;

  return true;
end;
$$;

grant execute on function public.set_utente_ruolo(text, text) to anon, authenticated;
