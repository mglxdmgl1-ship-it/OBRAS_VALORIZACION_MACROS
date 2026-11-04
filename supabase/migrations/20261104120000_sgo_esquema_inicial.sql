-- SGO Costos, presupuestos y valorizaciones: esquema en Supabase.
-- Tablas propias con prefijo sgo_. No modifica public.perfiles ni otras tablas del proyecto
-- (el proyecto y su lista de usuarios de auth se comparten con otra aplicación).
--
-- Roles (public.sgo_usuarios.rol):
--   admin   registra y edita todas las obras, crea usuarios y asigna accesos.
--   cliente solo consulta (y exporta) las obras que el administrador le asigna.

create table public.sgo_usuarios (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null,
  nombre text not null default '',
  rol text not null default 'cliente' check (rol in ('admin', 'cliente')),
  activo boolean not null default true,
  debe_cambiar_clave boolean not null default false,
  creado timestamptz not null default now(),
  creado_por uuid references auth.users(id) on delete set null
);
comment on table public.sgo_usuarios is 'SGO: usuarios con acceso al sistema. admin = registra y edita todo; cliente = solo consulta las obras asignadas.';

create table public.sgo_obras (
  id text primary key check (id ~ '^[A-Za-z0-9._-]{3,80}$'),
  nombre text not null default '',
  cui text not null default '',
  datos jsonb not null check (jsonb_typeof(datos) = 'object'),
  version integer not null default 1,
  creado timestamptz not null default now(),
  creado_por uuid references auth.users(id) on delete set null,
  actualizado timestamptz not null default now(),
  actualizado_por uuid references auth.users(id) on delete set null
);
comment on table public.sgo_obras is 'SGO: una fila por obra; datos = obra completa (presupuesto, A.C.U., metrados, almacén, equipos, valorizaciones).';

create table public.sgo_accesos (
  obra_id text not null references public.sgo_obras(id) on delete cascade,
  usuario_id uuid not null references public.sgo_usuarios(id) on delete cascade,
  creado timestamptz not null default now(),
  primary key (obra_id, usuario_id)
);
create index sgo_accesos_usuario_idx on public.sgo_accesos (usuario_id);
comment on table public.sgo_accesos is 'SGO: obras que puede consultar cada cliente.';

create table public.sgo_respaldos (
  id bigint generated always as identity primary key,
  obra_id text not null references public.sgo_obras(id) on delete cascade,
  version integer not null,
  datos jsonb not null,
  creado timestamptz not null default now(),
  creado_por uuid references auth.users(id) on delete set null
);
create index sgo_respaldos_obra_idx on public.sgo_respaldos (obra_id, creado desc);
comment on table public.sgo_respaldos is 'SGO: copia automática de cada obra (una por día; se guardan las 30 últimas).';

-- ------------------------------------------------------------------ funciones de permisos
create or replace function public.sgo_es_admin() returns boolean
language sql stable security definer set search_path = '' as $$
  select exists (select 1 from public.sgo_usuarios where id = (select auth.uid()) and rol = 'admin' and activo);
$$;

create or replace function public.sgo_puede_ver(p_obra text) returns boolean
language sql stable security definer set search_path = '' as $$
  select public.sgo_es_admin() or exists (
    select 1 from public.sgo_accesos a join public.sgo_usuarios u on u.id = a.usuario_id
    where a.obra_id = p_obra and a.usuario_id = (select auth.uid()) and u.activo);
$$;

-- el usuario marca que ya cambió la contraseña temporal
create or replace function public.sgo_clave_cambiada() returns void
language sql security definer set search_path = '' as $$
  update public.sgo_usuarios set debe_cambiar_clave = false where id = (select auth.uid());
$$;

-- solo para la función sgo-usuarios (service_role): ubica una cuenta existente por correo
create or replace function public.sgo_id_por_correo(p_email text) returns uuid
language sql stable security definer set search_path = '' as $$
  select id from auth.users where lower(email) = lower(trim(p_email)) limit 1;
$$;

-- ------------------------------------------------------------------ disparadores
-- versión (control de cambios simultáneos), nombre y CUI tomados de los datos, quién y cuándo
create or replace function public.sgo_obras_antes() returns trigger
language plpgsql set search_path = '' as $$
begin
  if tg_op = 'UPDATE' then
    new.id := old.id;
    new.version := old.version + 1;
    new.creado := old.creado;
    new.creado_por := old.creado_por;
  else
    new.version := 1;
    new.creado := now();
    new.creado_por := coalesce(new.creado_por, auth.uid());
  end if;
  new.datos := jsonb_set(new.datos, '{id}', to_jsonb(new.id));
  new.nombre := coalesce(nullif(trim(new.datos #>> '{datos,nombre_corto}'), ''), nullif(trim(new.datos #>> '{datos,nombre}'), ''), 'SIN NOMBRE');
  new.cui := coalesce(new.datos #>> '{datos,cui}', '');
  new.actualizado := now();
  new.actualizado_por := auth.uid();
  return new;
end $$;
create trigger sgo_obras_antes before insert or update on public.sgo_obras
  for each row execute function public.sgo_obras_antes();

-- respaldo automático: una copia por día (la versión anterior al primer cambio del día)
create or replace function public.sgo_respaldo_diario() returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if not exists (select 1 from public.sgo_respaldos where obra_id = old.id and creado > now() - interval '20 hours') then
    insert into public.sgo_respaldos (obra_id, version, datos, creado_por) values (old.id, old.version, old.datos, old.actualizado_por);
    delete from public.sgo_respaldos r where r.obra_id = old.id
      and r.id not in (select x.id from public.sgo_respaldos x where x.obra_id = old.id order by x.creado desc limit 30);
  end if;
  return null;
end $$;
create trigger sgo_obras_respaldo after update on public.sgo_obras
  for each row when (old.datos is distinct from new.datos) execute function public.sgo_respaldo_diario();

-- siempre debe quedar un administrador activo
create or replace function public.sgo_usuarios_guardia() returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if old.rol = 'admin' and old.activo and (tg_op = 'DELETE' or new.rol <> 'admin' or not new.activo) then
    if not exists (select 1 from public.sgo_usuarios where rol = 'admin' and activo and id <> old.id) then
      raise exception 'Debe quedar al menos un administrador activo.' using errcode = 'P0001';
    end if;
  end if;
  if tg_op = 'DELETE' then return old; end if;
  new.id := old.id;
  return new;
end $$;
create trigger sgo_usuarios_guardia before update or delete on public.sgo_usuarios
  for each row execute function public.sgo_usuarios_guardia();

-- ------------------------------------------------------------------ seguridad a nivel de fila
alter table public.sgo_usuarios enable row level security;
alter table public.sgo_obras enable row level security;
alter table public.sgo_accesos enable row level security;
alter table public.sgo_respaldos enable row level security;

create policy sgo_usuarios_ver on public.sgo_usuarios for select to authenticated
  using (id = (select auth.uid()) or (select public.sgo_es_admin()));
create policy sgo_usuarios_editar on public.sgo_usuarios for update to authenticated
  using ((select public.sgo_es_admin())) with check ((select public.sgo_es_admin()));

create policy sgo_obras_ver on public.sgo_obras for select to authenticated
  using (public.sgo_puede_ver(id));
create policy sgo_obras_crear on public.sgo_obras for insert to authenticated
  with check ((select public.sgo_es_admin()));
create policy sgo_obras_editar on public.sgo_obras for update to authenticated
  using ((select public.sgo_es_admin())) with check ((select public.sgo_es_admin()));
create policy sgo_obras_borrar on public.sgo_obras for delete to authenticated
  using ((select public.sgo_es_admin()));

create policy sgo_accesos_ver on public.sgo_accesos for select to authenticated
  using (usuario_id = (select auth.uid()) or (select public.sgo_es_admin()));
create policy sgo_accesos_crear on public.sgo_accesos for insert to authenticated
  with check ((select public.sgo_es_admin()));
create policy sgo_accesos_borrar on public.sgo_accesos for delete to authenticated
  using ((select public.sgo_es_admin()));

create policy sgo_respaldos_ver on public.sgo_respaldos for select to authenticated
  using ((select public.sgo_es_admin()));

-- ------------------------------------------------------------------ privilegios
revoke all on table public.sgo_usuarios, public.sgo_obras, public.sgo_accesos, public.sgo_respaldos from anon, authenticated;
grant select on table public.sgo_usuarios to authenticated;
grant update (nombre, rol, activo, debe_cambiar_clave) on table public.sgo_usuarios to authenticated;
grant select, delete on table public.sgo_obras to authenticated;
grant insert (id, datos) on table public.sgo_obras to authenticated;
grant update (datos) on table public.sgo_obras to authenticated;
grant select, insert, delete on table public.sgo_accesos to authenticated;
grant select on table public.sgo_respaldos to authenticated;
grant all on table public.sgo_usuarios, public.sgo_obras, public.sgo_accesos, public.sgo_respaldos to service_role;

revoke execute on function public.sgo_es_admin(), public.sgo_puede_ver(text), public.sgo_clave_cambiada(), public.sgo_id_por_correo(text),
  public.sgo_obras_antes(), public.sgo_respaldo_diario(), public.sgo_usuarios_guardia() from public, anon, authenticated;
grant execute on function public.sgo_es_admin(), public.sgo_puede_ver(text), public.sgo_clave_cambiada() to authenticated;
grant execute on function public.sgo_id_por_correo(text) to service_role;
