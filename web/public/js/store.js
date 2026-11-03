// Persistencia local (IndexedDB) de las obras. Cada obra se guarda completa como un objeto.
import { uid, hoy, clonar } from './util.js';
import { crearMotor } from './calc.js';

const DB = 'sgo-valorizaciones';
const VER = 1;
let dbp = null;

function abrir() {
  if (dbp) return dbp;
  dbp = new Promise((res, rej) => {
    const r = indexedDB.open(DB, VER);
    r.onupgradeneeded = () => {
      const db = r.result;
      if (!db.objectStoreNames.contains('obras')) db.createObjectStore('obras', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('cfg')) db.createObjectStore('cfg');
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

export const db = {
  listar: () => tx('obras', 'readonly', (s) => s.getAll()),
  leer: (id) => tx('obras', 'readonly', (s) => s.get(id)),
  guardar: (obra) => tx('obras', 'readwrite', (s) => s.put(obra)),
  borrar: (id) => tx('obras', 'readwrite', (s) => s.delete(id)),
  cfg: (k) => tx('cfg', 'readonly', (s) => s.get(k)),
  setCfg: (k, v) => tx('cfg', 'readwrite', (s) => s.put(v, k)),
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
  o.id = o.id || uid('obra-');
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
  obra: null,
  motor: null,
  obras: [],
  sucios: false,
  escuchas: new Set(),
};

let tGuardar = null;
/** Marca la obra como modificada: rehace el motor y guarda en 600 ms. */
export function cambio({ guardarYa = false } = {}) {
  if (!S.obra) return;
  S.obra.actualizado = new Date().toISOString();
  S.motor = crearMotor(S.obra);
  S.sucios = true;
  S.escuchas.forEach((f) => f('sucio'));
  clearTimeout(tGuardar);
  tGuardar = setTimeout(guardarAhora, guardarYa ? 0 : 600);
}
export async function guardarAhora() {
  clearTimeout(tGuardar);
  if (!S.obra) return;
  await db.guardar(S.obra);
  S.sucios = false;
  const i = S.obras.findIndex((x) => x.id === S.obra.id);
  const res = resumenObra(S.obra);
  if (i >= 0) S.obras[i] = res; else S.obras.push(res);
  S.escuchas.forEach((f) => f('guardado'));
}
export function resumenObra(o) {
  return { id: o.id, nombre: o.datos?.nombre_corto || o.datos?.nombre || 'SIN NOMBRE', cui: o.datos?.cui || '', actualizado: o.actualizado };
}
export async function cargarLista() {
  const todas = await db.listar();
  S.obras = todas.map(resumenObra).sort((a, b) => (a.actualizado < b.actualizado ? 1 : -1));
  return S.obras;
}
export async function abrirObra(id) {
  await guardarAhora();
  const o = await db.leer(id);
  if (!o) return null;
  S.obra = normalizar(o);
  S.motor = crearMotor(S.obra);
  await db.setCfg('activa', id);
  S.escuchas.forEach((f) => f('obra'));
  return S.obra;
}
export async function agregarObra(o, abrirla = true) {
  normalizar(o);
  if (S.obras.some((x) => x.id === o.id)) o.id = uid('obra-');
  await db.guardar(o);
  await cargarLista();
  if (abrirla) await abrirObra(o.id);
  return o;
}
export async function eliminarObra(id) {
  await db.borrar(id);
  if (S.obra?.id === id) { S.obra = null; S.motor = null; }
  await cargarLista();
}
export function copiaObra(o) {
  const c = clonar(o);
  c.id = uid('obra-');
  c.datos.nombre_corto = (c.datos.nombre_corto || c.datos.nombre) + ' (COPIA)';
  return c;
}
