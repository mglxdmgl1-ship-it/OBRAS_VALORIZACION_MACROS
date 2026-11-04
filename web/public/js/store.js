// Estado de la aplicación y guardado de las obras.
//   modo 'local': todo en IndexedDB del navegador (sin inicio de sesión).
//   modo 'nube' : Supabase (tabla sgo_obras) + copia en IndexedDB para trabajar sin conexión.
//                 Cada guardado sube la obra completa con control de versión (si otra persona guardó
//                 antes, se pregunta qué hacer). Los clientes son de solo lectura.
import { uid, hoy, clonar } from './util.js';
import { crearMotor } from './calc.js';
import * as nube from './nube.js';

const DB = 'sgo-valorizaciones';
const VER = 2;
let dbp = null;

function abrir() {
  if (dbp) return dbp;
  dbp = new Promise((res, rej) => {
    const r = indexedDB.open(DB, VER);
    r.onupgradeneeded = () => {
      const db = r.result;
      if (!db.objectStoreNames.contains('obras')) db.createObjectStore('obras', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('cfg')) db.createObjectStore('cfg');
      if (!db.objectStoreNames.contains('cache')) db.createObjectStore('cache', { keyPath: 'clave' });
    };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  return dbp;
}
async function tx(store, modo, fn) {
  const db = await abrir();
  return new Promise((res, rej) => {
    const t = db.transaction(store, modo);
    const s = t.objectStore(store);
    let out;
    const r = fn(s);
    if (r) r.onsuccess = () => (out = r.result);
    t.oncomplete = () => res(out);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error || new Error('Transacción cancelada'));
  });
}

/** IndexedDB: 'obras' (modo local), 'cache' (copias del modo nube), 'cfg' (preferencias). */
export const db = {
  listar: () => tx('obras', 'readonly', (s) => s.getAll()),
  leer: (id) => tx('obras', 'readonly', (s) => s.get(id)),
  guardar: (obra) => tx('obras', 'readwrite', (s) => s.put(obra)),
  borrar: (id) => tx('obras', 'readwrite', (s) => s.delete(id)),
  cfg: (k) => tx('cfg', 'readonly', (s) => s.get(k)),
  setCfg: (k, v) => tx('cfg', 'readwrite', (s) => s.put(v, k)),
  cacheLeer: (clave) => tx('cache', 'readonly', (s) => s.get(clave)),
  cacheGuardar: (reg) => tx('cache', 'readwrite', (s) => s.put(reg)),
  cacheBorrar: (clave) => tx('cache', 'readwrite', (s) => s.delete(clave)),
  cacheVaciar: () => tx('cache', 'readwrite', (s) => s.clear()),
  cacheListar: () => tx('cache', 'readonly', (s) => s.getAll()),
};

/** Plantilla de una obra vacía. */
export function obraVacia(nombre = 'NUEVA OBRA') {
  return normalizar({
    id: uid('obra-'),
    datos: { nombre, nombre_corto: nombre, cui: '', meta: '', entidad: '', dependencia: '', modalidad: 'ADMINISTRACIÓN DIRECTA', fecha_inicio: hoy(), plazo_dias: 120, ampliacion_dias: 0 },
    indirectos: [
      { concepto: 'GASTOS GENERALES DE EJECUCIÓN', monto: 0 },
      { concepto: 'GASTOS DE SUPERVISIÓN', monto: 0 },
      { concepto: 'GASTOS DE LIQUIDACIÓN', monto: 0 },
    ],
  });
}

/** Completa los campos que falten (obras antiguas o importadas). */
export function normalizar(o) {
  o.version = 1;
  o.id = String(o.id || uid('obra-')).replace(/[^A-Za-z0-9._-]/g, '-').slice(0, 80);
  if (o.id.length < 3) o.id = uid('obra-');
  o.datos = Object.assign({ nombre: '', nombre_corto: '', cui: '', meta: '', entidad: '', dependencia: '', modalidad: 'ADMINISTRACIÓN DIRECTA', fuente_financiamiento: '', region: '', provincia: '', distrito: '', localidad: '', fecha_inicio: hoy(), plazo_dias: 0, ampliacion_dias: 0 }, o.datos || {});
  o.parametros = Object.assign({ umbral: 0.2, jornada: 8, uso_max_hh: 0.85, no_almacen: 'AGUA PUESTA EN OBRA', combustibles: 'DIESEL;PETROLEO;GASOHOL;GASOLINA;KEROSENE', redondeo: 'arriba', dias_laborables: [1, 2, 3, 4, 5, 6] }, o.parametros || {});
  for (const k of ['responsables', 'indirectos', 'documentos', 'feriados', 'partidas', 'insumos', 'almacen', 'movimientos', 'equipos', 'partes', 'periodos']) if (!Array.isArray(o[k])) o[k] = [];
  for (const k of ['acu', 'metrados', 'detalles', 'anterior_inicial']) if (!o[k] || typeof o[k] !== 'object') o[k] = {};
  o.firmas = Array.isArray(o.firmas) ? o.firmas : ['RESIDENTE DE OBRA', 'SUPERVISOR DE OBRA', ''];
  o.cronograma = o.cronograma && Array.isArray(o.cronograma.meses) ? o.cronograma : { meses: [], montos: {} };
  o.cronograma.montos = o.cronograma.montos || {};
  o.financiero = o.financiero || { devengado_acumulado: {} };
  o.financiero.devengado_acumulado = o.financiero.devengado_acumulado || {};
  o.movimientos.forEach((m) => { if (!m.id) m.id = uid('M'); });
  o.partes.forEach((p) => { if (!p.id) p.id = uid('P'); });
  o.periodos.forEach((p) => { p.manual = p.manual || {}; });
  o.actualizado = o.actualizado || new Date().toISOString();
  return o;
}

// ------------------------------------------------------------------ estado de la aplicación
export const S = {
  modo: 'local',        // 'local' | 'nube'
  usuario: null,        // {id, email, nombre, rol, ...} en modo nube
  soloLectura: false,   // clientes
  obra: null,
  motor: null,
  obras: [],
  version: null,        // versión de la obra en el servidor
  sucios: false,        // cambios sin guardar en el navegador
  pendiente: false,     // cambios sin subir a la nube
  sync: 'ok',           // ok | guardando | pendiente | sin-conexion | conflicto | error
  errorSync: '',
  escuchas: new Set(),
};
/** Ganchos de la interfaz (los pone app.js): preguntar qué hacer ante un conflicto de versiones. */
export const ganchos = {
  conflicto: async () => 'servidor',
  recargado: () => {},
};
const avisar = (ev) => S.escuchas.forEach((f) => { try { f(ev); } catch (e) { console.error(e); } });
const claveCache = (id) => `${S.usuario?.id || 'anon'}|${id}`;
export const esAdmin = () => S.modo === 'local' || S.usuario?.rol === 'admin';
export const hayPendiente = () => S.sucios || S.pendiente;

export function configurarSesion(usuario) {
  S.modo = usuario ? 'nube' : 'local';
  S.usuario = usuario || null;
  S.soloLectura = !!usuario && usuario.rol !== 'admin';
  S.obra = null; S.motor = null; S.obras = []; S.version = null;
  S.sucios = false; S.pendiente = false; S.sync = 'ok';
}

let tLocal = null, tSubir = null, tReintento = null;
let nCambios = 0;
/** Marca la obra como modificada: rehace el motor, guarda en el navegador y sube a la nube. */
export function cambio({ guardarYa = false } = {}) {
  if (!S.obra) return;
  S.motor = crearMotor(S.obra);
  if (S.soloLectura) return; // los clientes no guardan
  S.obra.actualizado = new Date().toISOString();
  S.sucios = true;
  nCambios++;
  if (S.modo === 'nube') { S.pendiente = true; S.sync = 'pendiente'; }
  avisar('sucio');
  clearTimeout(tLocal);
  tLocal = setTimeout(guardarLocal, guardarYa ? 0 : 500);
  if (S.modo === 'nube') {
    clearTimeout(tSubir);
    tSubir = setTimeout(subir, guardarYa ? 50 : 2500);
  }
}
async function guardarLocal() {
  clearTimeout(tLocal);
  if (!S.obra || S.soloLectura) return;
  try {
    if (S.modo === 'local') {
      await db.guardar(S.obra);
      S.sucios = false;
      actualizarResumen();
      avisar('guardado');
    } else {
      await db.cacheGuardar({ clave: claveCache(S.obra.id), uid: S.usuario.id, id: S.obra.id, obra: S.obra, base: S.version, pendiente: S.pendiente, fecha: new Date().toISOString() });
      S.sucios = false;
    }
  } catch (e) {
    console.error(e);
    S.errorSync = 'No se pudo guardar en el navegador: ' + e.message;
    avisar('error');
  }
}
function actualizarResumen() {
  const i = S.obras.findIndex((x) => x.id === S.obra.id);
  const res = resumenObra(S.obra);
  if (S.modo === 'nube') res.version = S.version;
  if (i >= 0) S.obras[i] = { ...S.obras[i], ...res }; else S.obras.unshift(res);
}

let subiendo = null;
/** Sube la obra a Supabase (control de versión). */
export async function subir() {
  clearTimeout(tSubir);
  if (S.modo !== 'nube' || !S.obra || S.soloLectura || !S.pendiente) return;
  if (subiendo) { await subiendo; if (S.pendiente) return subir(); return; }
  const obra = S.obra, n0 = nCambios;
  subiendo = (async () => {
    S.sync = 'guardando'; avisar('sync');
    try {
      const r = await nube.guardarObra(obra, S.version);
      if (r.conflicto) { await resolverConflicto(); return; }
      if (S.obra !== obra) return;
      S.version = r.version;
      if (nCambios === n0) { S.pendiente = false; S.sync = 'ok'; }
      else S.sync = 'pendiente';
      S.errorSync = '';
      await guardarLocal();
      actualizarResumen();
      avisar('guardado');
      if (S.pendiente) tSubir = setTimeout(subir, 1500);
    } catch (e) {
      console.error(e);
      const sinRed = /conexión|internet|fetch|network/i.test(e.message);
      S.sync = sinRed ? 'sin-conexion' : 'error';
      S.errorSync = e.message;
      avisar('sync');
      clearTimeout(tReintento);
      tReintento = setTimeout(subir, sinRed ? 20000 : 60000);
    }
  })();
  try { await subiendo; } finally { subiendo = null; }
}
async function resolverConflicto() {
  const id = S.obra.id;
  const remoto = await nube.leerObra(id);
  if (!remoto) { S.sync = 'error'; S.errorSync = 'La obra ya no existe en el servidor (la eliminaron).'; avisar('sync'); return; }
  S.sync = 'conflicto'; avisar('sync');
  const decision = await ganchos.conflicto({ remoto, local: S.obra, versionLocal: S.version });
  if (decision === 'mio') {
    S.version = remoto.version;
    S.sync = 'pendiente';
    tSubir = setTimeout(subir, 10);
  } else {
    S.obra = normalizar(remoto.datos);
    S.version = remoto.version;
    S.motor = crearMotor(S.obra);
    S.pendiente = false; S.sucios = false; S.sync = 'ok';
    await db.cacheGuardar({ clave: claveCache(id), uid: S.usuario.id, id, obra: S.obra, base: S.version, pendiente: false, fecha: new Date().toISOString() });
    ganchos.recargado('conflicto');
  }
  avisar('sync');
}

/** Guarda todo lo pendiente (antes de cambiar de obra, de salir o de cerrar la página). */
export async function guardarAhora() {
  clearTimeout(tLocal);
  if (!S.obra || S.soloLectura) return;
  if (S.sucios) await guardarLocal();
  if (S.modo === 'nube' && S.pendiente) await subir();
}

export function resumenObra(o) {
  return { id: o.id, nombre: o.datos?.nombre_corto || o.datos?.nombre || 'SIN NOMBRE', cui: o.datos?.cui || '', actualizado: o.actualizado };
}
export async function cargarLista() {
  if (S.modo === 'nube') {
    try {
      const filas = await nube.listarObras();
      S.obras = filas.map((f) => ({ id: f.id, nombre: f.nombre, cui: f.cui, actualizado: f.actualizado, version: f.version, actualizado_por: f.actualizado_por }));
    } catch (e) {
      if (!/conexión|internet/i.test(e.message)) throw e;
      // sin conexión: las copias guardadas en este equipo
      const copias = (await db.cacheListar().catch(() => [])).filter((c) => c.uid === S.usuario?.id);
      S.obras = copias.map((c) => ({ ...resumenObra(c.obra), version: c.base }));
      S.sync = 'sin-conexion';
    }
    return S.obras;
  }
  const todas = await db.listar();
  S.obras = todas.map(resumenObra).sort((a, b) => (a.actualizado < b.actualizado ? 1 : -1));
  return S.obras;
}
const cfgActiva = () => (S.modo === 'nube' ? 'activa:' + S.usuario.id : 'activa');
export const obraActivaGuardada = () => db.cfg(cfgActiva());

/** Lee una obra completa (sin abrirla). */
export async function leerObraCompleta(id) {
  if (S.modo === 'nube') { const r = await nube.leerObra(id); return r ? r.datos : null; }
  return db.leer(id);
}

export async function abrirObra(id) {
  await guardarAhora();
  let obra = null, version = null, sinConexion = false;
  if (S.modo === 'nube') {
    let remoto = null, errRemoto = null;
    try { remoto = await nube.leerObra(id); } catch (e) { errRemoto = e; }
    const cache = await db.cacheLeer(claveCache(id)).catch(() => null);
    if (remoto) {
      if (cache?.pendiente && !S.soloLectura) {
        if (cache.base === remoto.version) { obra = cache.obra; version = remoto.version; S.pendiente = true; }
        else {
          const d = await ganchos.conflicto({ remoto, local: cache.obra, versionLocal: cache.base, alAbrir: true });
          if (d === 'mio') { obra = cache.obra; version = remoto.version; S.pendiente = true; }
          else { obra = remoto.datos; version = remoto.version; S.pendiente = false; }
        }
      } else { obra = remoto.datos; version = remoto.version; S.pendiente = false; }
    } else if (cache && errRemoto) {
      obra = cache.obra; version = cache.base; S.pendiente = !!cache.pendiente; sinConexion = true;
    } else if (errRemoto) throw errRemoto;
    else return null;
  } else {
    obra = await db.leer(id);
    if (!obra) return null;
  }
  S.obra = normalizar(obra);
  S.version = version;
  S.motor = crearMotor(S.obra);
  S.sucios = false;
  S.sync = sinConexion ? 'sin-conexion' : S.pendiente ? 'pendiente' : 'ok';
  await db.setCfg(cfgActiva(), id).catch(() => {});
  if (S.modo === 'nube' && !S.soloLectura) {
    await db.cacheGuardar({ clave: claveCache(id), uid: S.usuario.id, id, obra: S.obra, base: S.version, pendiente: S.pendiente, fecha: new Date().toISOString() }).catch(() => {});
    if (S.pendiente && !sinConexion) tSubir = setTimeout(subir, 300);
  }
  avisar('obra');
  return S.obra;
}

export async function agregarObra(o, abrirla = true) {
  normalizar(o);
  if (S.modo === 'nube') {
    if (S.soloLectura) throw new Error('Su cuenta es de solo lectura.');
    for (let i = 0; i < 3; i++) {
      try { await nube.crearObra(o); break; } catch (e) {
        if (e.codigo === '23505' && i < 2) { o.id = uid('obra-'); continue; }
        throw e;
      }
    }
  } else {
    if (S.obras.some((x) => x.id === o.id)) o.id = uid('obra-');
    await db.guardar(o);
  }
  await cargarLista();
  if (abrirla) await abrirObra(o.id);
  return o;
}
export async function eliminarObra(id) {
  if (S.modo === 'nube') {
    await nube.borrarObra(id);
    await db.cacheBorrar(claveCache(id)).catch(() => {});
  } else await db.borrar(id);
  if (S.obra?.id === id) { S.obra = null; S.motor = null; S.version = null; S.pendiente = false; S.sucios = false; }
  await cargarLista();
}
export function copiaObra(o) {
  const c = clonar(o);
  c.id = uid('obra-');
  c.datos.nombre_corto = (c.datos.nombre_corto || c.datos.nombre) + ' (COPIA)';
  return c;
}
/** Revisa si otra persona guardó una versión más nueva y la carga (si aquí no hay cambios sin subir). */
export async function revisarVersionRemota() {
  if (S.modo !== 'nube' || !S.obra || hayPendiente() || subiendo) return false;
  const r = await nube.versionObra(S.obra.id);
  if (!r || r.version === S.version) return false;
  const remoto = await nube.leerObra(S.obra.id);
  if (!remoto || hayPendiente()) return false;
  S.obra = normalizar(remoto.datos);
  S.version = remoto.version;
  S.motor = crearMotor(S.obra);
  if (!S.soloLectura) await db.cacheGuardar({ clave: claveCache(S.obra.id), uid: S.usuario.id, id: S.obra.id, obra: S.obra, base: S.version, pendiente: false, fecha: new Date().toISOString() }).catch(() => {});
  actualizarResumen();
  avisar('obra');
  ganchos.recargado('remoto');
  return true;
}
/** Al cerrar sesión: sube lo pendiente si se puede y borra las copias de este navegador. */
export async function cerrarSesionLocal() {
  try { await guardarAhora(); } catch { /* sin conexión */ }
  await db.cacheVaciar().catch(() => {});
  configurarSesion(null);
}
