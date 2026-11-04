// SGO – gestión de usuarios (solo administradores de SGO).
// Acciones (POST JSON):
//   { accion: 'crear', email, nombre, rol: 'admin'|'cliente', clave?, restablecer? }
//       Crea la cuenta (contraseña temporal, debe cambiarla al entrar) o, si el correo ya existe
//       en el proyecto (cuenta compartida con otra aplicación), solo le da acceso a SGO.
//   { accion: 'clave', id, clave }   Restablece la contraseña (temporal) de un usuario de SGO.
//   { accion: 'quitar', id }         Quita el acceso a SGO (la cuenta de auth no se borra).
// El rol, el nombre y el estado activo se cambian directamente en public.sgo_usuarios (RLS).
import { createClient } from 'npm:@supabase/supabase-js@2.45.4';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};
const json = (cuerpo: unknown, estado = 200) =>
  new Response(JSON.stringify(cuerpo), { status: estado, headers: { ...CORS, 'Content-Type': 'application/json; charset=utf-8' } });
const correoValido = (s: string) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s);

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CORS });
  if (req.method !== 'POST') return json({ error: 'Método no permitido.' }, 405);

  const url = Deno.env.get('SUPABASE_URL');
  const servicio = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY');
  if (!url || !servicio) return json({ error: 'La función no tiene configuradas las variables del proyecto.' }, 500);
  const sb = createClient(url, servicio, { auth: { persistSession: false, autoRefreshToken: false } });

  // ---- quién llama: debe ser administrador activo de SGO
  const token = (req.headers.get('Authorization') || '').replace(/^Bearer\s+/i, '').trim();
  if (!token) return json({ error: 'Inicie sesión.' }, 401);
  const { data: quien, error: eQuien } = await sb.auth.getUser(token);
  if (eQuien || !quien?.user) return json({ error: 'La sesión no es válida o venció. Vuelva a ingresar.' }, 401);
  const yo = quien.user;
  const { data: perfil } = await sb.from('sgo_usuarios').select('rol, activo').eq('id', yo.id).maybeSingle();
  if (!perfil || perfil.rol !== 'admin' || !perfil.activo) return json({ error: 'Solo un administrador de SGO puede gestionar usuarios.' }, 403);

  let b: Record<string, unknown>;
  try { b = await req.json(); } catch { return json({ error: 'Solicitud no válida.' }, 400); }
  const accion = String(b.accion || '');

  try {
    if (accion === 'crear') {
      const email = String(b.email || '').trim().toLowerCase();
      const nombre = String(b.nombre || '').trim().slice(0, 120);
      const rol = b.rol === 'admin' ? 'admin' : 'cliente';
      const clave = String(b.clave || '');
      const restablecer = b.restablecer === true;
      if (!correoValido(email)) return json({ error: 'El correo no es válido.' }, 400);
      if (clave && clave.length < 8) return json({ error: 'La contraseña debe tener al menos 8 caracteres.' }, 400);

      const { data: existente, error: eBus } = await sb.rpc('sgo_id_por_correo', { p_email: email });
      if (eBus) return json({ error: eBus.message }, 500);
      let id = (existente as string | null) || null;
      const existia = !!id;
      let temporal = false;
      if (!id) {
        if (!clave) return json({ error: 'Indique una contraseña temporal (mínimo 8 caracteres).' }, 400);
        const { data: nuevo, error: eNuevo } = await sb.auth.admin.createUser({ email, password: clave, email_confirm: true, user_metadata: { nombre } });
        if (eNuevo || !nuevo?.user) return json({ error: eNuevo?.message || 'No se pudo crear la cuenta.' }, 400);
        id = nuevo.user.id;
        temporal = true;
      } else if (restablecer) {
        if (!clave) return json({ error: 'Indique la nueva contraseña temporal.' }, 400);
        const { error: eClave } = await sb.auth.admin.updateUserById(id, { password: clave });
        if (eClave) return json({ error: eClave.message }, 400);
        temporal = true;
      }
      const { data: previo } = await sb.from('sgo_usuarios').select('id').eq('id', id).maybeSingle();
      if (previo) {
        const { error: eAct } = await sb.from('sgo_usuarios').update({ nombre, rol, activo: true, ...(temporal ? { debe_cambiar_clave: true } : {}) }).eq('id', id);
        if (eAct) return json({ error: eAct.message }, 400);
      } else {
        const { error: eIns } = await sb.from('sgo_usuarios').insert({ id, email, nombre, rol, activo: true, debe_cambiar_clave: temporal, creado_por: yo.id });
        if (eIns) return json({ error: eIns.message }, 400);
      }
      return json({ ok: true, id, existia, yaEraUsuario: !!previo, temporal });
    }

    if (accion === 'clave') {
      const id = String(b.id || '');
      const clave = String(b.clave || '');
      if (clave.length < 8) return json({ error: 'La contraseña debe tener al menos 8 caracteres.' }, 400);
      const { data: u } = await sb.from('sgo_usuarios').select('id').eq('id', id).maybeSingle();
      if (!u) return json({ error: 'Ese usuario no pertenece a SGO.' }, 404);
      const { error } = await sb.auth.admin.updateUserById(id, { password: clave });
      if (error) return json({ error: error.message }, 400);
      if (id !== yo.id) await sb.from('sgo_usuarios').update({ debe_cambiar_clave: true }).eq('id', id);
      return json({ ok: true });
    }

    if (accion === 'quitar') {
      const id = String(b.id || '');
      if (id === yo.id) return json({ error: 'No puede quitarse el acceso a usted mismo.' }, 400);
      const { error } = await sb.from('sgo_usuarios').delete().eq('id', id);
      if (error) return json({ error: error.message }, 400);
      return json({ ok: true });
    }

    return json({ error: 'Acción desconocida.' }, 400);
  } catch (e) {
    return json({ error: e instanceof Error ? e.message : String(e) }, 500);
  }
});
