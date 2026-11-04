-- Administrador inicial de SGO: la cuenta dueña del proyecto (ya existe en auth.users).
-- Entra con su misma contraseña. Desde «Usuarios y accesos» crea las demás cuentas.
-- Para otra obra o proyecto, cambie el correo y vuelva a ejecutar.
insert into public.sgo_usuarios (id, email, nombre, rol, activo)
select u.id, u.email, coalesce(u.raw_user_meta_data->>'nombre', ''), 'admin', true
from auth.users u
where lower(u.email) = 'mgl.xdmgl1@gmail.com'
on conflict (id) do update set rol = 'admin', activo = true;
