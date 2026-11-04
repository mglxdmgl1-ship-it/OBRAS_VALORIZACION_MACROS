// Aplicación: inicio de sesión (Supabase), menú según el rol, selector de obra, enrutador por #hash,
// guardado y sincronización automáticos, modo solo lectura para clientes.
import { S, cargarLista, abrirObra, guardarAhora, obraActivaGuardada, configurarSesion, ganchos, hayPendiente, revisarVersionRemota, subir, cerrarSesionLocal, esAdmin } from './store.js';
import { icono, $, aviso, ventana, confirmar, descargar, nombreArchivo } from './ui.js';
import { esc, fFecha, hoy } from './util.js';
import { nubeConfigurada, sesionActual, miPerfil, alCambiarSesion, salir } from './nube.js';
import { pantallaIngreso, pantallaCambioClave, pantallaSinAcceso, pantallaCargando, ocultarAcceso } from './acceso.js';

// roles: 'admin' = solo administradores (y modo local). nube: solo con inicio de sesión.
const RUTAS = [
  { g: 'General' },
  { id: 'panel', t: 'Panel de la obra', i: 'panel', m: () => import('./vistas/panel.js') },
  { id: 'obra', t: 'Datos de la obra', i: 'obra', m: () => import('./vistas/obra.js') },
  { g: 'Costos y presupuestos' },
  { id: 'presupuesto', t: 'Presupuesto', i: 'presupuesto', m: () => import('./vistas/presupuesto.js') },
  { id: 'acu', t: 'Análisis de costos (A.C.U.)', i: 'acu', m: () => import('./vistas/acu.js') },
  { id: 'insumos', t: 'Insumos y precios', i: 'insumos', m: () => import('./vistas/insumos.js') },
  { id: 'calculadora', t: 'Calculadora de insumos', i: 'calc', admin: true, m: () => import('./vistas/calculadora.js') },
  { g: 'Control de obra' },
  { id: 'metrados', t: 'Metrados diarios', i: 'metrados', m: () => import('./vistas/metrados.js') },
  { id: 'almacen', t: 'Almacén', i: 'almacen', m: () => import('./vistas/almacen.js') },
  { id: 'equipos', t: 'Equipo mecánico', i: 'equipos', m: () => import('./vistas/equipos.js') },
  { g: 'Valorización' },
  { id: 'valorizacion', t: 'Valorización mensual', i: 'valorizacion', m: () => import('./vistas/valorizacion.js') },
  { id: 'cronograma', t: 'Cronograma valorizado', i: 'cronograma', m: () => import('./vistas/cronograma.js') },
  { id: 'reportes', t: 'Reportes PDF / Excel', i: 'reportes', m: () => import('./vistas/reportes.js') },
  { g: 'Sistema' },
  { id: 'obras', t: () => (S.soloLectura ? 'Mis obras' : 'Obras, importar y respaldo'), i: 'respaldo', sinObra: true, m: () => import('./vistas/obras.js') },
  { id: 'usuarios', t: 'Usuarios y accesos', i: 'usuarios', admin: true, nube: true, sinObra: true, m: () => import('./vistas/usuarios.js') },
  { id: 'cuenta', t: 'Mi cuenta', i: 'cuenta', nube: true, sinObra: true, m: () => import('./vistas/cuenta.js') },
];
const titulo = (r) => (typeof r.t === 'function' ? r.t() : r.t);
const permitida = (r) => !r.g && (!r.admin || esAdmin()) && (!r.nube || S.modo === 'nube');

function pintarMenu(activa) {
  const items = [];
  for (let i = 0; i < RUTAS.length; i++) {
    const r = RUTAS[i];
    if (r.g) {
      const hay = RUTAS.slice(i + 1).findIndex((x) => x.g);
      const grupo = RUTAS.slice(i + 1, hay < 0 ? undefined : i + 1 + hay);
      if (grupo.some(permitida)) items.push(`<div class="menu-grupo">${r.g}</div>`);
    } else if (permitida(r)) items.push(`<a class="menu-item ${r.id === activa ? 'activo' : ''}" href="#/${r.id}">${icono(r.i)}<span>${titulo(r)}</span></a>`);
  }
  $('#menu').innerHTML = items.join('');
}
function pintarSelector() {
  const sel = $('#sel-obra');
  if (!S.obras.length) { sel.innerHTML = '<option>— sin obras —</option>'; sel.disabled = true; return; }
  sel.disabled = false;
  sel.innerHTML = S.obras.map((o) => `<option value="${esc(o.id)}" ${S.obra?.id === o.id ? 'selected' : ''}>${esc(o.nombre)}${o.cui ? ' · CUI ' + esc(o.cui) : ''}</option>`).join('');
}
function pintarUsuario() {
  const u = S.usuario;
  const box = $('#usuario');
  if (!u) { box.innerHTML = ''; box.hidden = true; return; }
  box.hidden = false;
  box.innerHTML = `<a class="usuario-chip" href="#/cuenta" title="${esc(u.email)}">${icono('cuenta')}<span class="usuario-nombre">${esc(u.nombre || u.email)}</span><span class="badge ${u.rol === 'admin' ? 'badge-azul' : 'badge-alerta'}">${u.rol === 'admin' ? 'Administrador' : 'Cliente · solo lectura'}</span></a>
    <button class="btn btn-sm" id="b-salir" title="Cerrar sesión">${icono('salir')}<span class="solo-escritorio">Salir</span></button>`;
  box.querySelector('#b-salir').onclick = cerrarSesion;
  document.body.classList.toggle('solo-lectura', S.soloLectura);
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
  let r = RUTAS.find((x) => x.id === ruta && permitida(x)) || RUTAS[1];
  if (!S.obra && !r.sinObra) r = RUTAS.find((x) => x.id === 'obras');
  pintarMenu(r.id);
  $('#lateral').classList.remove('abierto');
  const cont = $('#vista');
  try {
    const mod = await r.m();
    if (id !== renderId) return;
    if (vistaActual?.destruir) vistaActual.destruir();
    vistaActual = mod;
    document.title = titulo(r) + ' · SGO Costos y Valorizaciones';
    await mod.render(cont, q);
    bloquearSiCorresponde();
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

// ------------------------------------------------------------------ solo lectura (clientes)
// Se desactivan los campos de edición y se ocultan los botones que modifican datos; quedan los filtros,
// las pestañas, la navegación y las exportaciones. La base de datos (RLS) igual rechaza cualquier cambio.
const BTN_LIBRES = '[data-exp],[data-tab],[data-pag],[data-ir],[data-modo],[data-libre],[data-acc="acu"],[data-kx],[data-abrir]';
function bloquear(raiz) {
  raiz.querySelectorAll('input, select, textarea').forEach((el) => {
    if (el.dataset.sl) return;
    el.dataset.sl = '1';
    if (el.type === 'search' || el.matches('[data-libre]') || el.closest('.filtros, [data-libre]')) return;
    el.disabled = true;
  });
  raiz.querySelectorAll('button, a.btn').forEach((b) => {
    if (b.dataset.sl) return;
    b.dataset.sl = '1';
    if (b.matches(BTN_LIBRES) || b.closest('.pestanas, .pag, [data-libre]')) return;
    if (b.tagName === 'A' && (b.getAttribute('href') || '').startsWith('#/')) return;
    b.hidden = true;
  });
}
let rafBloqueo = 0;
function bloquearSiCorresponde() {
  if (!S.soloLectura) return;
  cancelAnimationFrame(rafBloqueo);
  rafBloqueo = requestAnimationFrame(() => bloquear($('#vista')));
}

// ------------------------------------------------------------------ estado de guardado / sincronización
function estadoGuardado(ev) {
  const e = $('#estado-guardado');
  const hora = () => new Date().toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' });
  if (S.modo === 'local') {
    if (ev === 'sucio') { e.textContent = 'Guardando…'; e.className = 'estado-guardado sucio'; }
    if (ev === 'guardado') { e.textContent = 'Guardado en este equipo ' + hora(); e.className = 'estado-guardado ok'; pintarSelector(); }
  } else if (S.soloLectura) {
    e.textContent = 'Solo lectura'; e.className = 'estado-guardado';
  } else {
    const t = { pendiente: ['Cambios por subir…', 'sucio'], guardando: ['Subiendo a la nube…', 'sucio'], ok: ['En la nube ' + hora(), 'ok'], 'sin-conexion': ['Sin conexión: guardado en este equipo', 'error'], error: ['Error al subir: reintentando', 'error'], conflicto: ['Conflicto de versiones', 'error'] }[S.sync] || ['', ''];
    if (ev === 'sucio' && S.sync === 'pendiente') t[0] = 'Cambios por subir…';
    e.textContent = t[0];
    e.className = 'estado-guardado ' + t[1];
    e.title = S.errorSync || '';
    if (ev === 'guardado') pintarSelector();
  }
  if (ev === 'obra') pintarSelector();
}

// ------------------------------------------------------------------ conflictos y recargas
ganchos.conflicto = ({ remoto, local, alAbrir }) => new Promise((resolver) => {
  let r = 'servidor';
  const nombre = remoto?.datos?.datos?.nombre_corto || remoto?.id || '';
  ventana({
    titulo: 'Otra persona guardó cambios en esta obra', ancho: '620px',
    cuerpo: `<p>${alAbrir ? 'En este equipo hay cambios de <b>' + esc(nombre) + '</b> que no se llegaron a subir' : 'Mientras usted trabajaba, se guardó otra versión de <b>' + esc(nombre) + '</b>'} y en el servidor ya hay una versión más nueva (N° ${remoto?.version}, del ${fFecha(String(remoto?.actualizado || '').slice(0, 10))} ${String(remoto?.actualizado || '').slice(11, 16)}).</p>
      <ul class="lista-simple"><li><b>Usar la versión del servidor</b>: se descartan sus cambios no subidos.</li><li><b>Conservar mis cambios</b>: su versión reemplaza a la del servidor (la del servidor queda en los respaldos automáticos del día).</li></ul>
      <p><button class="btn btn-sm" id="cf-desc">${icono('respaldo')}Descargar mi versión (.json) antes de decidir</button></p>`,
    botones: [{ txt: 'Usar la versión del servidor', accion: () => { r = 'servidor'; } }, { txt: 'Conservar mis cambios', clase: 'btn-primario', accion: () => { r = 'mio'; } }],
    alAbrir: (raiz) => raiz.querySelector('#cf-desc').addEventListener('click', () => descargar(nombreArchivo(`MI_VERSION_${local?.datos?.nombre_corto || 'obra'}_${hoy()}`) + '.json', JSON.stringify({ tipo: 'sgo-respaldo', fecha: new Date().toISOString(), obra: local }), 'application/json')),
    alCerrar: () => resolver(r),
  });
});
ganchos.recargado = (motivo) => {
  aviso(motivo === 'remoto' ? 'Se cargó la versión más reciente de la obra (guardada por otra persona).' : 'Se cargó la versión del servidor.', 'ok', 5000);
  refrescar();
};

// ------------------------------------------------------------------ sesión
async function cerrarSesion() {
  if (S.modo !== 'nube') return;
  try { await guardarAhora(); } catch { /* sin conexión */ }
  if (hayPendiente() && !(await confirmar('Hay cambios que todavía no se subieron a la nube (sin conexión). Si cierra la sesión ahora, <b>se perderán</b>. ¿Cerrar sesión de todas formas?', { peligro: true, si: 'Cerrar sesión' }))) return;
  await cerrarSesionLocal();
  try { localStorage.removeItem('sgo-perfil'); } catch { /* sin almacenamiento */ }
  await salir();
  location.hash = '';
  location.reload();
}
/** Perfil de SGO del usuario de la sesión (con copia para abrir sin conexión). */
async function cargarPerfil(ses) {
  try {
    const p = await miPerfil(ses.user.id);
    try { if (p) localStorage.setItem('sgo-perfil', JSON.stringify(p)); else localStorage.removeItem('sgo-perfil'); } catch { /* */ }
    return { perfil: p };
  } catch (e) {
    let guardado = null;
    try { guardado = JSON.parse(localStorage.getItem('sgo-perfil') || 'null'); } catch { /* */ }
    if (guardado && guardado.id === ses.user.id && /conexión|internet/i.test(e.message)) return { perfil: guardado, sinConexion: true };
    return { error: e.message };
  }
}
async function flujoNube() {
  pantallaCargando();
  let ses = null;
  try { ses = await sesionActual(); } catch (e) { console.warn(e); }
  if (!ses) ses = await pantallaIngreso();
  pantallaCargando('Cargando su cuenta…');
  const { perfil, error, sinConexion } = await cargarPerfil(ses);
  if (error) { pantallaSinAcceso({ error }); return false; }
  if (!perfil || !perfil.activo) { pantallaSinAcceso({ email: ses.user.email, desactivado: !!perfil }); return false; }
  if (perfil.debe_cambiar_clave && !sinConexion) {
    await pantallaCambioClave({ obligatorio: true, email: perfil.email });
    perfil.debe_cambiar_clave = false;
  }
  configurarSesion(perfil);
  ocultarAcceso();
  pintarUsuario();
  if (sinConexion) aviso('Sin conexión con el servidor: se muestran las copias guardadas en este equipo.', 'alerta', 8000);
  alCambiarSesion((ev) => { if (ev === 'SIGNED_OUT') { location.hash = ''; location.reload(); } });
  return true;
}

async function cargarDatos() {
  try {
    await cargarLista();
    const activa = await obraActivaGuardada().catch(() => null);
    const id = S.obras.find((o) => o.id === activa)?.id || S.obras[0]?.id;
    if (id) await abrirObra(id);
  } catch (e) {
    console.error(e);
    aviso((S.modo === 'nube' ? 'No se pudieron cargar las obras: ' : 'No se pudo abrir la base de datos local del navegador: ') + e.message, 'error', 10000);
  }
  pintarSelector();
}

async function iniciar() {
  S.escuchas.add(estadoGuardado);
  $('#sel-obra').addEventListener('change', async (e) => {
    try {
      await abrirObra(e.target.value);
      aviso('Obra activa: ' + (S.obra?.datos?.nombre_corto || ''));
    } catch (err) { aviso('No se pudo abrir la obra: ' + err.message, 'error', 8000); }
    render();
  });
  $('#btn-menu').addEventListener('click', () => $('#lateral').classList.toggle('abierto'));
  window.addEventListener('hashchange', render);
  window.addEventListener('beforeunload', (e) => { if (hayPendiente()) { guardarAhora(); e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && hayPendiente()) guardarAhora(); });
  new MutationObserver(() => bloquearSiCorresponde()).observe($('#vista'), { childList: true, subtree: true });

  if (nubeConfigurada()) {
    if (!(await flujoNube())) return;
    window.addEventListener('online', () => subir());
    // si otra persona guardó una versión más nueva (y aquí no hay cambios sin subir), se carga sola
    setInterval(async () => {
      if (document.visibilityState !== 'visible') return;
      try { await revisarVersionRemota(); } catch { /* sin conexión */ }
    }, 60000);
  } else {
    configurarSesion(null);
    ocultarAcceso();
  }
  await cargarDatos();
  estadoGuardado('inicio');
  render();
}
iniciar();
