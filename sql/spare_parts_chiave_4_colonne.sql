-- Compatibilità (spare_parts): la riga non è più identificata da (pn, terminal_pn) ma da
-- (pn, terminal_pn, pnit, type).
--
-- Motivo: lo stesso PN può comparire sullo stesso Terminal PN con SPARE diversi (es. 301601020000382
-- su Q58: COM-CONNETTORE RS232 + COM-LAN CONNECTOR) o su PNIT diversi (es. A77-3TA040 / A77-3TA948);
-- con la vecchia chiave l'import teneva solo la prima riga e scartava le altre come duplicati.
--
-- Da eseguire una volta nel SQL Editor di Supabase, PRIMA di usare la nuova versione dell'app.
-- Lo script è idempotente. Dopo l'esecuzione, reimportare il file Excel di Compatibilità.

-- 1) pnit e type fanno parte della chiave: niente NULL (in un vincolo UNIQUE i NULL non si confrontano).
update public.spare_parts set pnit = coalesce(pnit, ''), type = coalesce(type, '')
where pnit is null or type is null;

alter table public.spare_parts alter column pnit set default '';
alter table public.spare_parts alter column pnit set not null;
alter table public.spare_parts alter column type set default '';
alter table public.spare_parts alter column type set not null;

-- 2) Elimina il vecchio vincolo/indice univoco su (pn, terminal_pn), qualunque nome abbia.
do $$
declare
  r record;
begin
  for r in
    select i.indexrelid::regclass::text as idx, c.conname
    from pg_index i
    left join pg_constraint c on c.conindid = i.indexrelid and c.contype in ('u', 'p')
    where i.indrelid = 'public.spare_parts'::regclass
      and i.indisunique
      and not i.indisprimary
      and (
        select array_agg(a.attname::text order by a.attname)
        from pg_attribute a
        where a.attrelid = i.indrelid and a.attnum = any (i.indkey)
      ) = array['pn', 'terminal_pn']
  loop
    if r.conname is not null then
      execute format('alter table public.spare_parts drop constraint %I', r.conname);
    else
      execute format('drop index %s', r.idx);
    end if;
  end loop;
end $$;

-- 3) Nuova chiave. Serve a PostgREST per l'upsert con on_conflict=pn,terminal_pn,pnit,type.
do $$
begin
  if not exists (
    select 1 from pg_constraint
    where conrelid = 'public.spare_parts'::regclass and conname = 'spare_parts_pn_terminal_pnit_type_key'
  ) then
    alter table public.spare_parts
      add constraint spare_parts_pn_terminal_pnit_type_key unique (pn, terminal_pn, pnit, type);
  end if;
end $$;
