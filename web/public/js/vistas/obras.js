// Gestión de obras: crear, abrir, duplicar, eliminar, respaldo e importación de JSON.
import { S, agregarObra, abrirObra, eliminarObra, obraVacia, copiaObra, db, normalizar } from '../store.js';
import { esc, fFecha, hoy } from '../util.js';
import { icono, aviso, confirmar, descargar, elegirArchivo, nombreArchivo, formulario, trabajando } from '../ui.js';
import { obraDesdeValorizacionJSON } from '../importar.js';
import { periodoInicial } from '../calc.js';
import { cab } from './comun.js';
import { refrescar, ir } from '../app.js';

export async function render(cont) {
  const hay = S.obras.length > 0;
  cont.innerHTML = `
    ${cab('Obras, importar y respaldo', 'Cada obra guarda su presupuesto, A.C.U., metrados diarios, almacén, equipos y valorizaciones. Los datos quedan en este navegador: descargue un respaldo (.json) con frecuencia y para pasar la obra a otra computadora.')}
    ${hay ? '' : `<div class="panel"><div class="panel-cuerpo"><h2 style="color:var(--azul);margin-bottom:6px">Bienvenido</h2><p class="tenue" style="margin:0">Para empezar, cargue la obra de ejemplo (I.E.I. N° 881, con su presupuesto, 3 600 filas de A.C.U., metrados de octubre 2026, almacén y la valorización N° 01) o cree una obra nueva y llene los datos a mano o importándolos desde Excel.</p></div></div>`}
    <div class="rep-lista" style="margin-bottom:18px">
      <div class="rep"><h4>${icono('mas')} Nueva obra en blanco</h4><p>Empiece de cero: datos generales, luego digite o importe el presupuesto, los A.C.U. y el cronograma.</p><div class="acciones"><button class="btn btn-primario" id="b-nueva">${icono('mas')}Crear obra</button></div></div>
      <div class="rep"><h4>${icono('obra')} Obra de ejemplo: I.E.I. N° 881</h4><p>Datos reales de los libros de Excel del repositorio (Cabanilla – Lampa – Puno). Útil para conocer el programa.</p><div class="acciones"><button class="btn btn-oro" id="b-ejemplo">${icono('importar')}Cargar ejemplo</button></div></div>
      <div class="rep"><h4>${icono('importar')} Importar respaldo o valorización</h4><p>Respaldo completo (.json descargado de este programa) o el JSON de la valorización exportado con <span class="mono">exportar_valorizacion_json.py</span>.</p><div class="acciones"><button class="btn" id="b-importar">${icono('importar')}Elegir archivo .json</button></div></div>
    </div>
    ${hay ? `<div class="panel"><div class="panel-cab"><h2>Obras registradas</h2><div class="acciones"><button class="btn" id="b-respaldo-todo">${icono('respaldo')}Respaldo de todas</button></div></div>
      <div class="tabla-cont libre"><table class="t"><thead><tr><th>Obra</th><th>CUI</th><th>Última modificación</th><th></th></tr></thead><tbody>
      ${S.obras.map((o) => `<tr class="${S.obra?.id === o.id ? 'sel' : ''}"><td><b>${esc(o.nombre)}</b>${S.obra?.id === o.id ? ' <span class="badge badge-azul">activa</span>' : ''}</td><td class="mono">${esc(o.cui)}</td><td>${fFecha(o.actualizado?.slice(0, 10))} ${o.actualizado?.slice(11, 16) || ''}</td>
        <td class="acc"><button class="btn btn-sm" data-abrir="${o.id}">${icono('ojo')}Abrir</button> <button class="btn btn-sm" data-resp="${o.id}">${icono('respaldo')}Respaldo</button> <button class="btn btn-sm" data-dup="${o.id}">${icono('copiar')}Duplicar</button> <button class="btn-icono peligro" data-del="${o.id}" title="Eliminar">${icono('borrar')}</button></td></tr>`).join('')}
      </tbody></table></div></div>` : ''}
    <div class="panel"><div class="panel-cab"><h2>Orden de trabajo recomendado</h2></div><div class="panel-cuerpo">
      ${[
        ['Datos de la obra', 'Entidad, proyecto, CUI, meta, fecha de inicio, plazo, responsables y costos indirectos.', 'obra'],
        ['Presupuesto', 'Importe el presupuesto de S10/Delphin (Excel) o digítelo. Títulos y partidas con metrado y P.U.', 'presupuesto'],
        ['A.C.U. e insumos', 'Importe los análisis de costos unitarios o arme cada A.C.U. Cambie precios de insumos y se actualizan todos los análisis.', 'acu'],
        ['Cronograma valorizado', 'Importe o genere el programado por mes (para el cuadro 4.5 y la curva S).', 'cronograma'],
        ['Cada día: metrados, almacén y equipos', 'Digite el metrado del día; el programa calcula los insumos (A.C.U. × metrado), los estados (VÁLIDO, NO VALORIZAR…) y propone las salidas del almacén y los partes de equipo, que puede redondear y editar.', 'metrados'],
        ['Fin de mes: valorización', 'Revise la valorización 4.4, el resumen 4.3, el comparado 4.5 y la curva S; exporte a PDF/Excel y cierre el periodo.', 'valorizacion'],
      ].map(([t, d, r], i) => `<div class="paso"><div class="paso-n">${i + 1}</div><div><h4><a href="#/${r}">${t}</a></h4><p>${d}</p></div></div>`).join('')}
    </div></div>`;

  cont.querySelector('#b-nueva').onclick = () => nuevaObra();
  cont.querySelector('#b-ejemplo').onclick = () => cargarEjemplo();
  cont.querySelector('#b-importar').onclick = () => importarJSON();
  cont.querySelector('#b-respaldo-todo')?.addEventListener('click', async () => {
    const todas = await db.listar();
    descargar(`RESPALDO_TODAS_LAS_OBRAS_${hoy()}.json`, JSON.stringify({ tipo: 'sgo-respaldo-multiple', fecha: new Date().toISOString(), obras: todas }), 'application/json');
  });
  cont.querySelectorAll('[data-abrir]').forEach((b) => (b.onclick = async () => { await abrirObra(b.dataset.abrir); ir('panel'); }));
  cont.querySelectorAll('[data-resp]').forEach((b) => (b.onclick = async () => {
    const o = await db.leer(b.dataset.resp);
    descargar(nombreArchivo(`RESPALDO_${o.datos.nombre_corto || o.datos.nombre}_${hoy()}`) + '.json', JSON.stringify({ tipo: 'sgo-respaldo', fecha: new Date().toISOString(), obra: o }), 'application/json');
  }));
  cont.querySelectorAll('[data-dup]').forEach((b) => (b.onclick = async () => {
    const o = await db.leer(b.dataset.dup);
    await agregarObra(copiaObra(o));
    aviso('Obra duplicada.');
    refrescar();
  }));
  cont.querySelectorAll('[data-del]').forEach((b) => (b.onclick = async () => {
    const o = S.obras.find((x) => x.id === b.dataset.del);
    if (!(await confirmar(`¿Eliminar la obra <b>${esc(o.nombre)}</b> y todos sus datos? Esta acción no se puede deshacer. Descargue antes un respaldo si lo necesita.`, { peligro: true, si: 'Eliminar' }))) return;
    await eliminarObra(o.id);
    if (!S.obra && S.obras.length) await abrirObra(S.obras[0].id);
    aviso('Obra eliminada.');
    refrescar();
  }));
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
  await trabajando('Cargando la obra de ejemplo…', async () => {
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
    for (const o of obras) {
      normalizar(o);
      if (S.obras.some((x) => x.id === o.id)) {
        const reemplazar = await confirmar(`Ya existe la obra «${esc(o.datos.nombre_corto || o.datos.nombre)}». ¿Reemplazarla con el respaldo? (Cancelar = importarla como copia)`, { si: 'Reemplazar' });
        if (!reemplazar) o.id = o.id + '-' + Date.now().toString(36);
      }
      await db.guardar(o);
    }
    const { cargarLista } = await import('../store.js');
    await cargarLista();
    await abrirObra(obras[obras.length - 1].id);
    aviso(`${obras.length} obra(s) importada(s).`);
    ir('panel');
  } catch (e) {
    aviso('No se pudo importar: ' + e.message, 'error', 8000);
  }
}
