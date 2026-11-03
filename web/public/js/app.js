// Aplicación: menú, selector de obra, enrutador por #hash y guardado automático.
import { S, cargarLista, abrirObra, db, guardarAhora } from './store.js';
import { icono, $, aviso } from './ui.js';
import { esc } from './util.js';

const RUTAS = [
  { g: 'General' },
  { id: 'panel', t: 'Panel de la obra', i: 'panel', m: () => import('./vistas/panel.js') },
  { id: 'obra', t: 'Datos de la obra', i: 'obra', m: () => import('./vistas/obra.js') },
  { g: 'Costos y presupuestos' },
  { id: 'presupuesto', t: 'Presupuesto', i: 'presupuesto', m: () => import('./vistas/presupuesto.js') },
  { id: 'acu', t: 'Análisis de costos (A.C.U.)', i: 'acu', m: () => import('./vistas/acu.js') },
  { id: 'insumos', t: 'Insumos y precios', i: 'insumos', m: () => import('./vistas/insumos.js') },
  { id: 'calculadora', t: 'Calculadora de insumos', i: 'calc', m: () => import('./vistas/calculadora.js') },
  { g: 'Control de obra' },
  { id: 'metrados', t: 'Metrados diarios', i: 'metrados', m: () => import('./vistas/metrados.js') },
  { id: 'almacen', t: 'Almacén', i: 'almacen', m: () => import('./vistas/almacen.js') },
  { id: 'equipos', t: 'Equipo mecánico', i: 'equipos', m: () => import('./vistas/equipos.js') },
  { g: 'Valorización' },
  { id: 'valorizacion', t: 'Valorización mensual', i: 'valorizacion', m: () => import('./vistas/valorizacion.js') },
  { id: 'cronograma', t: 'Cronograma valorizado', i: 'cronograma', m: () => import('./vistas/cronograma.js') },
  { id: 'reportes', t: 'Reportes PDF / Excel', i: 'reportes', m: () => import('./vistas/reportes.js') },
  { g: 'Sistema' },
  { id: 'obras', t: 'Obras, importar y respaldo', i: 'respaldo', m: () => import('./vistas/obras.js') },
];

function pintarMenu(activa) {
  $('#menu').innerHTML = RUTAS.map((r) => (r.g ? `<div class="menu-grupo">${r.g}</div>` : `<a class="menu-item ${r.id === activa ? 'activo' : ''}" href="#/${r.id}">${icono(r.i)}<span>${r.t}</span></a>`)).join('');
}
function pintarSelector() {
  const sel = $('#sel-obra');
  if (!S.obras.length) { sel.innerHTML = '<option>— sin obras —</option>'; sel.disabled = true; return; }
  sel.disabled = false;
  sel.innerHTML = S.obras.map((o) => `<option value="${esc(o.id)}" ${S.obra?.id === o.id ? 'selected' : ''}>${esc(o.nombre)}${o.cui ? ' · CUI ' + esc(o.cui) : ''}</option>`).join('');
}

export function parametros() {
  const h = location.hash.replace(/^#\/?/, '');
  const [ruta, qs] = h.split('?');
  return { ruta: ruta || 'panel', q: Object.fromEntries(new URLSearchParams(qs || '')) };
}
export function ir(ruta, q = {}) {
  const s = new URLSearchParams(q).toString();
  location.hash = '#/' + ruta + (s ? '?' + s : '');
}

let vistaActual = null;
let renderId = 0;
export async function render() {
  const id = ++renderId;
  const { ruta, q } = parametros();
  let r = RUTAS.find((x) => x.id === ruta) || RUTAS[1];
  if (!S.obra && r.id !== 'obras') r = RUTAS.find((x) => x.id === 'obras');
  pintarMenu(r.id);
  $('#lateral').classList.remove('abierto');
  const cont = $('#vista');
  try {
    const mod = await r.m();
    if (id !== renderId) return;
    if (vistaActual?.destruir) vistaActual.destruir();
    vistaActual = mod;
    document.title = r.t + ' · SGO Costos y Valorizaciones';
    await mod.render(cont, q);
    // los parámetros (?item=…, ?dia=…) se usan una sola vez
    if (Object.keys(q).length) history.replaceState(null, '', '#/' + r.id);
  } catch (e) {
    console.error(e);
    cont.innerHTML = `<div class="nota nota-error"><b>Error al mostrar la pantalla.</b><br>${esc(e.message || e)}</div>`;
  }
}
/** Vuelve a pintar la vista actual (después de un cambio). */
export function refrescar() {
  return render();
}

function estadoGuardado(ev) {
  const e = $('#estado-guardado');
  if (ev === 'sucio') { e.textContent = 'Guardando…'; e.className = 'estado-guardado sucio'; }
  if (ev === 'guardado') { e.textContent = 'Guardado ' + new Date().toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' }); e.className = 'estado-guardado ok'; pintarSelector(); }
  if (ev === 'obra') pintarSelector();
}

async function iniciar() {
  S.escuchas.add(estadoGuardado);
  try {
    await cargarLista();
    const activa = await db.cfg('activa');
    const id = S.obras.find((o) => o.id === activa)?.id || S.obras[0]?.id;
    if (id) await abrirObra(id);
  } catch (e) {
    console.error(e);
    aviso('No se pudo abrir la base de datos local del navegador: ' + e.message, 'error', 10000);
  }
  pintarSelector();
  $('#sel-obra').addEventListener('change', async (e) => {
    await abrirObra(e.target.value);
    aviso('Obra activa: ' + (S.obra?.datos?.nombre_corto || ''));
    render();
  });
  $('#btn-menu').addEventListener('click', () => $('#lateral').classList.toggle('abierto'));
  window.addEventListener('hashchange', render);
  window.addEventListener('beforeunload', (e) => { if (S.sucios) { guardarAhora(); e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && S.sucios) guardarAhora(); });
  render();
}
iniciar();
