// Gestión de obras: crear, abrir, duplicar, eliminar, respaldo e importación de JSON.
// En la nube: el administrador gestiona las obras y sus respaldos automáticos; el cliente solo ve «Mis obras».
import { S, agregarObra, abrirObra, eliminarObra, obraVacia, copiaObra, db, normalizar, leerObraCompleta, cargarLista, esAdmin, cambio } from '../store.js';
import { esc, fFecha, hoy } from '../util.js';
import { icono, aviso, confirmar, descargar, elegirArchivo, nombreArchivo, formulario, trabajando } from '../ui.js';
import { obraDesdeValorizacionJSON } from '../importar.js';
import { periodoInicial } from '../calc.js';
import * as nube from '../nube.js';
import { cab } from './comun.js';
import { refrescar, ir } from '../app.js';

const fechaHora = (s) => (s ? `${fFecha(String(s).slice(0, 10))} ${new Date(s).toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' })}` : '');

export async function render(cont) {
  if (S.modo === 'nube' && !esAdmin()) return misObras(cont);
  const hay = S.obras.length > 0;
  const nube_ = S.modo === 'nube';
  cont.innerHTML = `
    ${cab('Obras, importar y respaldo', nube_
      ? 'Las obras se guardan en la nube (Supabase) y se comparten con los usuarios autorizados. Cada día se guarda automáticamente una copia de respaldo de cada obra.'
      : 'Cada obra guarda su presupuesto, A.C.U., metrados diarios, almacén, equipos y valorizaciones. Los datos quedan en este navegador: descargue un respaldo (.json) con frecuencia y para pasar la obra a otra computadora.')}
    ${hay ? '' : `<div class="panel"><div class="panel-cuerpo"><h2 style="color:var(--azul);margin-bottom:6px">Bienvenido</h2><p class="tenue" style="margin:0">Para empezar, cargue la obra de ejemplo (I.E.I. N° 881, con su presupuesto, 3 600 filas de A.C.U., metrados de octubre 2026, almacén y la valorización N° 01) o cree una obra nueva y llene los datos a mano o importándolos desde Excel.</p></div></div>`}
    <div class="rep-lista" style="margin-bottom:18px">
      <div class="rep"><h4>${icono('mas')} Nueva obra en blanco</h4><p>Empiece de cero: datos generales, luego digite o importe el presupuesto, los A.C.U. y el cronograma.</p><div class="acciones"><button class="btn btn-primario" id="b-nueva">${icono('mas')}Crear obra</button></div></div>
      <div class="rep"><h4>${icono('obra')} Obra de ejemplo: I.E.I. N° 881</h4><p>Datos reales de los libros de Excel del repositorio (Cabanilla – Lampa – Puno). Útil para conocer el programa.</p><div class="acciones"><button class="btn btn-oro" id="b-ejemplo">${icono('importar')}Cargar ejemplo</button></div></div>
      <div class="rep"><h4>${icono('importar')} Importar respaldo o valorización</h4><p>Respaldo completo (.json descargado de este programa) o el JSON de la valorización exportado con <span class="mono">exportar_valorizacion_json.py</span>.</p><div class="acciones"><button class="btn" id="b-importar">${icono('importar')}Elegir archivo .json</button></div></div>
    </div>
    <div id="locales"></div>
    ${hay ? `<div class="panel"><div class="panel-cab"><h2>Obras registradas${nube_ ? ' en la nube' : ''}</h2><div class="acciones">${nube_ ? `<a class="btn btn-sm" href="#/usuarios">${icono('usuarios')}Usuarios y accesos</a>` : ''}<button class="btn btn-sm" id="b-respaldo-todo">${icono('respaldo')}Respaldo de todas</button></div></div>
      <div class="tabla-cont libre"><table class="t"><thead><tr><th>Obra</th><th>CUI</th><th>Última modificación</th>${nube_ ? '<th class="n">Versión</th>' : ''}<th></th></tr></thead><tbody>
      ${S.obras.map((o) => `<tr class="${S.obra?.id === o.id ? 'sel' : ''}"><td><b>${esc(o.nombre)}</b>${S.obra?.id === o.id ? ' <span class="badge badge-azul">activa</span>' : ''}</td><td class="mono">${esc(o.cui)}</td><td>${fechaHora(o.actualizado)}</td>${nube_ ? `<td class="n">${o.version ?? ''}</td>` : ''}
        <td class="acc"><button class="btn btn-sm" data-abrir="${esc(o.id)}">${icono('ojo')}Abrir</button> <button class="btn btn-sm" data-resp="${esc(o.id)}">${icono('respaldo')}Respaldo</button> <button class="btn btn-sm" data-dup="${esc(o.id)}">${icono('copiar')}Duplicar</button> <button class="btn-icono peligro" data-del="${esc(o.id)}" title="Eliminar">${icono('borrar')}</button></td></tr>`).join('')}
      </tbody></table></div></div>` : ''}
    ${nube_ && S.obra ? `<div class="panel"><div class="panel-cab"><h2>Respaldos automáticos de «${esc(S.obra.datos.nombre_corto || S.obra.datos.nombre)}»</h2><span class="tenue" style="font-size:12px">Una copia por día (las 30 últimas)</span></div><div class="panel-cuerpo" id="respaldos"><span class="tenue">Cargando…</span></div></div>` : ''}
    <div class="panel"><div class="panel-cab"><h2>Orden de trabajo recomendado</h2></div><div class="panel-cuerpo">
      ${[
        ['Datos de la obra', 'Entidad, proyecto, CUI, meta, fecha de inicio, plazo, responsables y costos indirectos.', 'obra'],
        ['Presupuesto', 'Importe el presupuesto de S10/Delphin (Excel) o digítelo. Títulos y partidas con metrado y P.U.', 'presupuesto'],
        ['A.C.U. e insumos', 'Importe los análisis de costos unitarios o arme cada A.C.U. Cambie precios de insumos y se actualizan todos los análisis.', 'acu'],
        ['Cronograma valorizado', 'Importe o genere el programado por mes (para el cuadro 4.5 y la curva S).', 'cronograma'],
        ['Cada día: metrados, almacén y equipos', 'Digite el metrado del día; el programa calcula los insumos (A.C.U. × metrado), los estados (VÁLIDO, NO VALORIZAR…) y propone las salidas del almacén y los partes de equipo, que puede redondear y editar.', 'metrados'],
        ['Fin de mes: valorización', 'Revise la valorización 4.4, el resumen 4.3, el comparado 4.5 y la curva S; exporte a PDF/Excel y cierre el periodo.', 'valorizacion'],
        ...(nube_ ? [['Clientes', 'Cree las cuentas de los clientes (entidad, supervisión) y asígneles sus obras: verán el avance y descargarán los reportes sin poder modificar nada.', 'usuarios']] : []),
      ].map(([t, d, r], i) => `<div class="paso"><div class="paso-n">${i + 1}</div><div><h4><a href="#/${r}">${t}</a></h4><p>${d}</p></div></div>`).join('')}
    </div></div>`;

  cont.querySelector('#b-nueva').onclick = () => nuevaObra();
  cont.querySelector('#b-ejemplo').onclick = () => cargarEjemplo().catch((e) => aviso(e.message, 'error', 8000));
  cont.querySelector('#b-importar').onclick = () => importarJSON();
  cont.querySelector('#b-respaldo-todo')?.addEventListener('click', async () => {
    try {
      const todas = await trabajando('Preparando el respaldo…', async () => {
        const out = [];
        for (const o of S.obras) out.push(await leerObraCompleta(o.id));
        return out.filter(Boolean);
      });
      descargar(`RESPALDO_TODAS_LAS_OBRAS_${hoy()}.json`, JSON.stringify({ tipo: 'sgo-respaldo-multiple', fecha: new Date().toISOString(), obras: todas }), 'application/json');
    } catch (e) { aviso(e.message, 'error', 8000); }
  });
  cont.querySelectorAll('[data-abrir]').forEach((b) => (b.onclick = async () => { try { await abrirObra(b.dataset.abrir); ir('panel'); } catch (e) { aviso(e.message, 'error', 8000); } }));
  cont.querySelectorAll('[data-resp]').forEach((b) => (b.onclick = async () => {
    try {
      const o = await leerObraCompleta(b.dataset.resp);
      descargar(nombreArchivo(`RESPALDO_${o.datos.nombre_corto || o.datos.nombre}_${hoy()}`) + '.json', JSON.stringify({ tipo: 'sgo-respaldo', fecha: new Date().toISOString(), obra: o }), 'application/json');
    } catch (e) { aviso(e.message, 'error', 8000); }
  }));
  cont.querySelectorAll('[data-dup]').forEach((b) => (b.onclick = async () => {
    try {
      const o = await leerObraCompleta(b.dataset.dup);
      await trabajando('Duplicando…', () => agregarObra(copiaObra(o)));
      aviso('Obra duplicada.');
      refrescar();
    } catch (e) { aviso(e.message, 'error', 8000); }
  }));
  cont.querySelectorAll('[data-del]').forEach((b) => (b.onclick = async () => {
    const o = S.obras.find((x) => x.id === b.dataset.del);
    if (!(await confirmar(`¿Eliminar la obra <b>${esc(o.nombre)}</b> y todos sus datos${nube_ ? ', sus respaldos automáticos y los accesos de los clientes' : ''}? Esta acción no se puede deshacer. Descargue antes un respaldo si lo necesita.`, { peligro: true, si: 'Eliminar' }))) return;
    try {
      await eliminarObra(o.id);
      if (!S.obra && S.obras.length) await abrirObra(S.obras[0].id);
      aviso('Obra eliminada.');
      refrescar();
    } catch (e) { aviso(e.message, 'error', 8000); }
  }));
  if (nube_) { pintarLocales(cont.querySelector('#locales')); if (S.obra) pintarRespaldos(cont.querySelector('#respaldos')); }
}

/** Vista del cliente: las obras que le asignaron. */
function misObras(cont) {
  cont.innerHTML = `
    ${cab('Mis obras', 'Obras que el administrador le asignó. Puede ver toda la información y descargar los reportes en PDF y Excel.')}
    ${S.obras.length ? `<div class="rep-lista">${S.obras.map((o) => `<div class="rep" ${S.obra?.id === o.id ? 'style="border-color:var(--oro)"' : ''}><h4>${icono('obra')} ${esc(o.nombre)}</h4><p>${o.cui ? 'CUI ' + esc(o.cui) + ' · ' : ''}actualizada el ${fechaHora(o.actualizado)}</p><div class="acciones"><button class="btn ${S.obra?.id === o.id ? '' : 'btn-primario'}" data-abrir="${esc(o.id)}">${icono('ojo')}${S.obra?.id === o.id ? 'Abierta' : 'Abrir'}</button></div></div>`).join('')}</div>`
      : `<div class="panel"><div class="vacio"><h3>Aún no tiene obras asignadas</h3><p>Pida al administrador del sistema que le dé acceso a sus obras.</p></div></div>`}`;
  cont.querySelectorAll('[data-abrir]').forEach((b) => (b.onclick = async () => { try { await abrirObra(b.dataset.abrir); ir('panel'); } catch (e) { aviso(e.message, 'error', 8000); } }));
}

/** Obras que quedaron solo en este navegador (de antes de usar la nube): subirlas. */
async function pintarLocales(box) {
  let locales = [];
  try { locales = await db.listar(); } catch { return; }
  const enNube = new Set(S.obras.map((o) => o.id));
  locales = locales.filter((o) => o?.datos);
  if (!locales.length) return;
  box.innerHTML = `<div class="panel"><div class="panel-cab"><h2>Obras guardadas solo en este navegador</h2><span class="tenue" style="font-size:12px">De cuando el programa trabajaba sin inicio de sesión</span></div><div class="panel-cuerpo">
    <table class="t"><tbody>${locales.map((o) => `<tr><td><b>${esc(o.datos.nombre_corto || o.datos.nombre)}</b> <span class="tenue">${esc(o.datos.cui || '')}</span></td><td>${fechaHora(o.actualizado)}</td><td class="acc">${enNube.has(o.id) ? '<span class="badge badge-ok">ya está en la nube</span>' : `<button class="btn btn-sm btn-primario" data-subir="${esc(o.id)}">${icono('nube')}Subir a la nube</button>`} <button class="btn-icono peligro" data-borrar-local="${esc(o.id)}" title="Borrar la copia de este navegador">${icono('borrar')}</button></td></tr>`).join('')}</tbody></table></div></div>`;
  box.querySelectorAll('[data-subir]').forEach((b) => (b.onclick = async () => {
    try {
      const o = locales.find((x) => x.id === b.dataset.subir);
      await trabajando('Subiendo a la nube…', () => agregarObra(JSON.parse(JSON.stringify(o)), false));
      aviso('Obra subida a la nube.');
      refrescar();
    } catch (e) { aviso(e.message, 'error', 8000); }
  }));
  box.querySelectorAll('[data-borrar-local]').forEach((b) => (b.onclick = async () => {
    if (!(await confirmar('¿Borrar la copia guardada en este navegador? (No afecta a la nube.)', { peligro: true }))) return;
    await db.borrar(b.dataset.borrarLocal);
    refrescar();
  }));
}

async function pintarRespaldos(box) {
  if (!box) return;
  try {
    const lista = await nube.listarRespaldos(S.obra.id);
    if (!lista.length) { box.innerHTML = '<span class="tenue">Todavía no hay respaldos: se crea el primero al guardar cambios en un nuevo día.</span>'; return; }
    box.innerHTML = `<table class="t"><thead><tr><th>Fecha</th><th class="n">Versión</th><th></th></tr></thead><tbody>${lista.map((r) => `<tr><td>${fechaHora(r.creado)}</td><td class="n">${r.version}</td><td class="acc"><button class="btn btn-sm" data-rdesc="${r.id}">${icono('respaldo')}Descargar</button> <button class="btn btn-sm" data-rrest="${r.id}">Restaurar</button></td></tr>`).join('')}</tbody></table>`;
    box.querySelectorAll('[data-rdesc]').forEach((b) => (b.onclick = async () => {
      try { const r = await nube.leerRespaldo(Number(b.dataset.rdesc)); descargar(nombreArchivo(`RESPALDO_${r.datos.datos?.nombre_corto || r.obra_id}_v${r.version}_${String(r.creado).slice(0, 10)}`) + '.json', JSON.stringify({ tipo: 'sgo-respaldo', fecha: r.creado, obra: r.datos }), 'application/json'); }
      catch (e) { aviso(e.message, 'error', 8000); }
    }));
    box.querySelectorAll('[data-rrest]').forEach((b) => (b.onclick = async () => {
      if (!(await confirmar('¿Restaurar este respaldo? La obra vuelve a como estaba en esa fecha (la versión actual queda como respaldo de hoy).', { peligro: true, si: 'Restaurar' }))) return;
      try {
        const r = await nube.leerRespaldo(Number(b.dataset.rrest));
        const datos = normalizar(r.datos);
        datos.id = S.obra.id;
        Object.keys(S.obra).forEach((k) => delete S.obra[k]);
        Object.assign(S.obra, datos);
        cambio({ guardarYa: true });
        aviso('Respaldo restaurado.');
        refrescar();
      } catch (e) { aviso(e.message, 'error', 8000); }
    }));
  } catch (e) { box.innerHTML = `<span class="neg">${esc(e.message)}</span>`; }
}

function nuevaObra() {
  const campos = [
    { id: 'nombre', et: 'Nombre del proyecto (como figura en el expediente)', req: true, col: 3 },
    { id: 'nombre_corto', et: 'Nombre corto', req: true, ph: 'I.E. N° 123 – DISTRITO' },
    { id: 'cui', et: 'CUI' }, { id: 'meta', et: 'Meta' },
    { id: 'entidad', et: 'Entidad / unidad ejecutora', col: 2, v: 'GOBIERNO REGIONAL DE PUNO' },
    { id: 'modalidad', et: 'Modalidad', tipo: 'select', ops: ['ADMINISTRACIÓN DIRECTA', 'CONTRATA', 'NÚCLEO EJECUTOR'] },
    { id: 'fecha_inicio', et: 'Fecha de inicio', tipo: 'date', v: hoy(), req: true },
    { id: 'plazo_dias', et: 'Plazo (días calendario)', tipo: 'number', v: 120, req: true },
  ];
  formulario({
    titulo: 'Nueva obra', campos, guardar: 'Crear obra',
    alGuardar: async (v) => {
      const o = obraVacia(v.nombre);
      Object.assign(o.datos, v);
      o.periodos = [periodoInicial(o)].filter(Boolean);
      await agregarObra(o);
      aviso('Obra creada. Ahora cargue el presupuesto.');
      ir('presupuesto');
    },
  });
}

export async function cargarEjemplo() {
  await trabajando(S.modo === 'nube' ? 'Cargando la obra de ejemplo en la nube…' : 'Cargando la obra de ejemplo…', async () => {
    const r = await fetch('datos/iei881.json', { cache: 'no-cache' });
    if (!r.ok) throw new Error('No se pudo descargar el ejemplo (' + r.status + ').');
    const o = await r.json();
    o.id = 'iei881-' + Date.now().toString(36);
    await agregarObra(o);
  });
  aviso('Obra de ejemplo cargada: I.E.I. N° 881.');
  ir('panel');
}

async function importarJSON() {
  const f = await elegirArchivo('.json,application/json');
  if (!f) return;
  try {
    const d = JSON.parse(await f.text());
    let obras = [];
    if (d.tipo === 'sgo-respaldo' && d.obra) obras = [d.obra];
    else if (d.tipo === 'sgo-respaldo-multiple') obras = d.obras || [];
    else if (d.partidas && d.datos) obras = [d];
    else if (d.valorizacion?.partidas) obras = [obraDesdeValorizacionJSON(d)];
    else throw new Error('No se reconoce el archivo. Use un respaldo de este programa o el JSON de la valorización.');
    let ultima = null;
    for (const o of obras) {
      normalizar(o);
      if (S.obras.some((x) => x.id === o.id)) {
        const reemplazar = await confirmar(`Ya existe la obra «${esc(o.datos.nombre_corto || o.datos.nombre)}». ¿Reemplazarla con el respaldo? (Cancelar = importarla como copia)`, { si: 'Reemplazar' });
        if (reemplazar) {
          if (S.modo === 'nube') {
            await abrirObra(o.id);
            Object.keys(S.obra).forEach((k) => delete S.obra[k]);
            Object.assign(S.obra, o);
            cambio({ guardarYa: true });
            ultima = o.id;
            continue;
          }
          await db.guardar(o);
          ultima = o.id;
          continue;
        }
        o.id = o.id + '-' + Date.now().toString(36);
      }
      await trabajando('Importando…', () => agregarObra(o, false));
      ultima = o.id;
    }
    await cargarLista();
    if (ultima) await abrirObra(ultima);
    aviso(`${obras.length} obra(s) importada(s).`);
    ir('panel');
  } catch (e) {
    aviso('No se pudo importar: ' + e.message, 'error', 8000);
  }
}
