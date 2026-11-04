// Conexión con Supabase: sesión, perfil y rol, obras (con control de versión), usuarios, accesos y respaldos.
// Tablas: public.sgo_usuarios, sgo_obras, sgo_accesos, sgo_respaldos (ver supabase/migrations).
import { SUPABASE } from './config.js';

let cli = null;
export const nubeConfigurada = () => !!(SUPABASE && SUPABASE.url && SUPABASE.clave);

export function cliente() {
  if (!cli) {
    if (!window.supabase?.createClient) throw new Error('No se cargó la librería de Supabase.');
    cli = window.supabase.createClient(SUPABASE.url, SUPABASE.clave, {
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: false, storageKey: 'sgo-sesion' },
    });
  }
  return cli;
}

/** Mensajes en español para los errores más comunes. */
export function traducir(e) {
  const m = String(e?.message || e || '');
  const c = String(e?.code || '');
  if (/Invalid login credentials/i.test(m)) return 'Correo o contraseña incorrectos.';
  if (/Email not confirmed/i.test(m)) return 'La cuenta aún no está confirmada. Pida al administrador que la active.';
  if (/rate limit|too many/i.test(m)) return 'Demasiados intentos. Espere unos minutos y vuelva a intentarlo.';
  if (/Failed to fetch|NetworkError|Load failed|network/i.test(m)) return 'Sin conexión con el servidor. Revise su internet.';
  if (/JWT expired|invalid JWT|session.*(missing|expired)/i.test(m)) return 'La sesión venció. Vuelva a ingresar.';
  if (/Password should be at least|password.*(short|length)/i.test(m)) return 'La contraseña es muy corta (mínimo 8 caracteres).';
  if (/same.*password|New password should be different/i.test(m)) return 'La nueva contraseña debe ser distinta de la actual.';
  if (/weak|pwned|leaked/i.test(m)) return 'La contraseña es muy débil o aparece en filtraciones conocidas. Use otra.';
  if (c === '42501' || /row-level security|permission denied/i.test(m)) return 'No tiene permiso para esta acción.';
  if (c === '23505' || /duplicate key/i.test(m)) return 'Ya existe un registro con ese código.';
  if (/al menos un administrador/i.test(m)) return 'Debe quedar al menos un administrador activo.';
  if (/relation .* does not exist|Could not find the table/i.test(m)) return 'La base de datos de SGO aún no está instalada en Supabase (faltan las tablas sgo_).';
  return m || 'Error desconocido.';
}
function falla(error) {
  const e = new Error(traducir(error));
  e.original = error;
  e.codigo = error?.code;
  return e;
}
const ok = ({ data, error }) => { if (error) throw falla(error); return data; };

// ------------------------------------------------------------------ sesión
export async function sesionActual() {
  const { data, error } = await cliente().auth.getSession();
  if (error) throw falla(error);
  return data.session;
}
export async function ingresar(email, clave) {
  const { data, error } = await cliente().auth.signInWithPassword({ email: String(email).trim(), password: clave });
  if (error) throw falla(error);
  return data.session;
}
export async function salir() {
  try { await cliente().auth.signOut(); } catch { /* sin conexión: igual se borra la sesión local */ }
}
export function alCambiarSesion(fn) {
  return cliente().auth.onAuthStateChange((ev, ses) => fn(ev, ses));
}
export async function cambiarMiClave(clave) {
  const { error } = await cliente().auth.updateUser({ password: clave });
  if (error) throw falla(error);
  const r = await cliente().rpc('sgo_clave_cambiada');
  if (r.error) throw falla(r.error);
}
export async function miPerfil(uid) {
  return ok(await cliente().from('sgo_usuarios').select('id, email, nombre, rol, activo, debe_cambiar_clave').eq('id', uid).maybeSingle());
}

// ------------------------------------------------------------------ obras
export async function listarObras() {
  return ok(await cliente().from('sgo_obras').select('id, nombre, cui, version, actualizado, actualizado_por').order('actualizado', { ascending: false }));
}
export async function leerObra(id) {
  return ok(await cliente().from('sgo_obras').select('id, datos, version, actualizado, actualizado_por').eq('id', id).maybeSingle());
}
export async function crearObra(obra) {
  return ok(await cliente().from('sgo_obras').insert({ id: obra.id, datos: obra }).select('id, version, actualizado').single());
}
/** Guarda si la versión del servidor sigue siendo `version`. Devuelve {version} o {conflicto:true}. */
export async function guardarObra(obra, version) {
  const filas = ok(await cliente().from('sgo_obras').update({ datos: obra }).eq('id', obra.id).eq('version', version).select('version, actualizado'));
  if (!filas || !filas.length) return { conflicto: true };
  return { version: filas[0].version, actualizado: filas[0].actualizado };
}
export async function versionObra(id) {
  return ok(await cliente().from('sgo_obras').select('version, actualizado, actualizado_por').eq('id', id).maybeSingle());
}
export async function borrarObra(id) {
  ok(await cliente().from('sgo_obras').delete().eq('id', id));
}

// ------------------------------------------------------------------ usuarios y accesos (administrador)
export async function listarUsuarios() {
  return ok(await cliente().from('sgo_usuarios').select('id, email, nombre, rol, activo, debe_cambiar_clave, creado').order('rol').order('nombre'));
}
export async function actualizarUsuario(id, campos) {
  return ok(await cliente().from('sgo_usuarios').update(campos).eq('id', id).select('id'));
}
/** Llama a la función sgo-usuarios (crear, clave, quitar). */
export async function funcionUsuarios(cuerpo) {
  const { data, error } = await cliente().functions.invoke(SUPABASE.funcionUsuarios || 'sgo-usuarios', { body: cuerpo });
  if (error) {
    let msg = error.message;
    try { const j = await error.context?.json?.(); if (j?.error) msg = j.error; } catch { /* sin cuerpo */ }
    if (/Failed to send a request|Failed to fetch/i.test(msg)) msg = 'No se pudo contactar la función de usuarios (sgo-usuarios). Verifique que esté publicada en Supabase.';
    throw new Error(msg);
  }
  if (data?.error) throw new Error(data.error);
  return data;
}
export async function listarAccesos() {
  return ok(await cliente().from('sgo_accesos').select('obra_id, usuario_id'));
}
export async function darAcceso(obra_id, usuario_id) {
  ok(await cliente().from('sgo_accesos').insert({ obra_id, usuario_id }));
}
export async function quitarAcceso(obra_id, usuario_id) {
  ok(await cliente().from('sgo_accesos').delete().eq('obra_id', obra_id).eq('usuario_id', usuario_id));
}
export async function listarRespaldos(obra_id) {
  return ok(await cliente().from('sgo_respaldos').select('id, version, creado, creado_por').eq('obra_id', obra_id).order('creado', { ascending: false }));
}
export async function leerRespaldo(id) {
  return ok(await cliente().from('sgo_respaldos').select('id, obra_id, version, datos, creado').eq('id', id).single());
}
