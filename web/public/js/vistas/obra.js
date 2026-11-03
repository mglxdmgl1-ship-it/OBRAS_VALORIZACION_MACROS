// Datos de la obra: identificación, plazo, responsables, costos indirectos, documentos, feriados y parámetros.
import { S, cambio } from '../store.js';
import { esc, fmt, pct, fFecha, num } from '../util.js';
import { icono, htmlCampos, leerCampos, aviso, confirmar } from '../ui.js';
import { fechaTermino } from '../calc.js';
import { cab } from './comun.js';
import { refrescar } from '../app.js';

const CAMPOS = [
  { tipo: 'seccion', et: 'Identificación' },
  { id: 'nombre', et: 'Nombre del proyecto', col: 3, req: true },
  { id: 'nombre_corto', et: 'Nombre corto', req: true },
  { id: 'cui', et: 'CUI' }, { id: 'meta', et: 'Meta' },
  { id: 'entidad', et: 'Entidad / unidad ejecutora', col: 2 }, { id: 'dependencia', et: 'Dependencia (gerencia / sub gerencia)' },
  { id: 'pliego', et: 'Pliego / U.E. (SIAF)', col: 2 },
  { id: 'modalidad', et: 'Modalidad de ejecución', tipo: 'select', ops: ['ADMINISTRACIÓN DIRECTA', 'CONTRATA', 'NÚCLEO EJECUTOR'] },
  { id: 'fuente_financiamiento', et: 'Fuente de financiamiento', col: 2 },
  { id: 'funcion', et: 'Función' }, { id: 'division_funcional', et: 'División funcional' }, { id: 'grupo_funcional', et: 'Grupo funcional' },
  { tipo: 'seccion', et: 'Ubicación' },
  { id: 'region', et: 'Región' }, { id: 'provincia', et: 'Provincia' }, { id: 'distrito', et: 'Distrito' }, { id: 'localidad', et: 'Localidad', col: 3 },
  { tipo: 'seccion', et: 'Plazo de ejecución' },
  { id: 'fecha_inicio', et: 'Fecha de inicio (acta de inicio)', tipo: 'date', req: true },
  { id: 'plazo_dias', et: 'Plazo (días calendario)', tipo: 'number', req: true },
  { id: 'ampliacion_dias', et: 'Ampliaciones de plazo (días)', tipo: 'number' },
];

export function render(cont) {
  const o = S.obra, m = S.motor;
  const t = fechaTermino(o.datos);
  const cd = m.costoDirecto();
  cont.innerHTML = `
    ${cab('Datos de la obra', 'Estos datos salen en el encabezado de todos los reportes (PDF y Excel).', `<button class="btn btn-primario" id="b-guardar">${icono('guardar')}Guardar datos</button>`)}
    <div class="panel"><div class="panel-cuerpo" id="f-datos">${htmlCampos(CAMPOS, o.datos)}
      <p class="nota">Fecha de término: <b>${fFecha(t) || '—'}</b> (inicio + plazo + ampliaciones − 1). Los días fuera del plazo se marcan «FUERA DEL PLAZO» y no se valorizan.</p></div></div>

    <div class="rejilla-2">
      <div class="panel"><div class="panel-cab"><h2>Responsables y firmas</h2><button class="btn btn-sm" id="b-resp">${icono('mas')}Agregar</button></div>
        <div class="panel-cuerpo">
          <table class="t"><thead><tr><th>Cargo</th><th>Nombre</th><th>Profesión</th><th></th></tr></thead><tbody id="t-resp">
          ${o.responsables.map((r, i) => `<tr><td><input data-r="${i}" data-k="cargo" value="${esc(r.cargo)}"></td><td><input data-r="${i}" data-k="nombre" value="${esc(r.nombre)}"></td><td><input data-r="${i}" data-k="profesion" value="${esc(r.profesion || '')}"></td><td><button class="btn-icono peligro" data-rdel="${i}">${icono('borrar')}</button></td></tr>`).join('') || '<tr><td colspan="4" class="tenue">Sin responsables</td></tr>'}
          </tbody></table>
          <div class="form-grid" style="margin-top:12px">
            ${[0, 1, 2].map((i) => `<label class="campo"><span>Firma ${i + 1} (cargo)</span><input data-firma="${i}" value="${esc(o.firmas[i] || '')}" placeholder="${['RESIDENTE DE OBRA', 'SUPERVISOR DE OBRA', 'opcional'][i]}"></label>`).join('')}
          </div>
          <p class="tenue" style="font-size:12px">El nombre de la firma se toma del responsable con el mismo cargo. No registre DNI ni celulares: la página puede ser pública.</p>
        </div></div>

      <div class="panel"><div class="panel-cab"><h2>Costos indirectos</h2><button class="btn btn-sm" id="b-ind">${icono('mas')}Agregar</button></div>
        <div class="panel-cuerpo">
          <table class="t"><thead><tr><th>Concepto</th><th class="n">Monto S/</th><th class="n">% del C.D.</th><th></th></tr></thead><tbody>
          <tr class="tit"><td>COSTO DIRECTO (del presupuesto)</td><td class="n">${fmt(cd)}</td><td class="n">100.00 %</td><td></td></tr>
          ${o.indirectos.map((x, i) => `<tr><td><input data-i="${i}" data-k="concepto" value="${esc(x.concepto)}"></td><td><input class="n" data-i="${i}" data-k="monto" value="${fmt(x.monto)}"></td><td class="n">${cd ? pct(x.monto / cd) : '—'}</td><td><button class="btn-icono peligro" data-idel="${i}">${icono('borrar')}</button></td></tr>`).join('')}
          <tr class="tot"><td>PRESUPUESTO TOTAL</td><td class="n">${fmt(m.presupuestoTotal())}</td><td class="n">${cd ? pct(m.presupuestoTotal() / cd) : ''}</td><td></td></tr>
          </tbody></table>
          <p class="tenue" style="font-size:12px">Cada indirecto se valoriza en proporción al C.D. valorizado: monto / C.D. × C.D. del mes. ¿Tiene el porcentaje? <button class="btn btn-sm" id="b-pct">Calcular montos desde %</button></p>
        </div></div>
    </div>

    <div class="rejilla-2">
      <div class="panel"><div class="panel-cab"><h2>Documentos (resoluciones, actas)</h2><button class="btn btn-sm" id="b-doc">${icono('mas')}Agregar</button></div>
        <div class="panel-cuerpo"><table class="t"><thead><tr><th>Documento</th><th>Resolución / N°</th><th>Fecha</th><th class="n">Monto</th><th></th></tr></thead><tbody>
        ${o.documentos.map((x, i) => `<tr><td><input data-d="${i}" data-k="documento" value="${esc(x.documento)}"></td><td><input data-d="${i}" data-k="resolucion" value="${esc(x.resolucion || '')}"></td><td><input type="date" data-d="${i}" data-k="fecha" value="${esc(x.fecha || '')}"></td><td><input class="n" data-d="${i}" data-k="monto" value="${x.monto ? fmt(x.monto) : ''}"></td><td><button class="btn-icono peligro" data-ddel="${i}">${icono('borrar')}</button></td></tr>`).join('') || '<tr><td colspan="5" class="tenue">Sin documentos</td></tr>'}
        </tbody></table></div></div>
      <div class="panel"><div class="panel-cab"><h2>Feriados y parámetros de control</h2><button class="btn btn-sm" id="b-fer">${icono('mas')}Feriado</button></div>
        <div class="panel-cuerpo">
          <table class="t"><thead><tr><th>Fecha</th><th>Descripción</th><th></th></tr></thead><tbody>
          ${o.feriados.map((x, i) => `<tr><td><input type="date" data-f="${i}" data-k="fecha" value="${esc(x.fecha)}"></td><td><input data-f="${i}" data-k="descripcion" value="${esc(x.descripcion)}"></td><td><button class="btn-icono peligro" data-fdel="${i}">${icono('borrar')}</button></td></tr>`).join('') || '<tr><td colspan="3" class="tenue">Sin feriados</td></tr>'}
          </tbody></table>
          <div class="form-grid" style="margin-top:14px">
            <label class="campo"><span>Umbral «NO VALORIZAR» (% del P.U.)</span><input data-par="umbral" value="${fmt((o.parametros.umbral ?? 0.2) * 100, 1)}"><small>Insumos sin O/S ni ingreso que pesen este % o más del P.U.</small></label>
            <label class="campo"><span>Jornada diaria (horas)</span><input data-par="jornada" value="${o.parametros.jornada ?? 8}"></label>
            <label class="campo"><span>Redondeo para el almacén</span><select data-par="redondeo">${[['arriba', 'Entero hacia arriba'], ['normal', 'Entero más cercano'], ['abajo', 'Entero hacia abajo'], ['dos', '2 decimales'], ['ninguno', 'Sin redondeo']].map(([k, t]) => `<option value="${k}" ${o.parametros.redondeo === k ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
            <label class="campo c3"><span>Insumos que no pasan por el almacén (palabras separadas por ;)</span><input data-par="no_almacen" value="${esc(o.parametros.no_almacen || '')}"></label>
          </div>
        </div></div>
    </div>`;

  const leerTodo = () => {
    Object.assign(o.datos, leerCampos(cont.querySelector('#f-datos'), CAMPOS));
    cont.querySelectorAll('[data-r]').forEach((e) => (o.responsables[e.dataset.r][e.dataset.k] = e.value.trim()));
    cont.querySelectorAll('[data-i]').forEach((e) => (o.indirectos[e.dataset.i][e.dataset.k] = e.dataset.k === 'monto' ? num(e.value) : e.value.trim()));
    cont.querySelectorAll('[data-d]').forEach((e) => (o.documentos[e.dataset.d][e.dataset.k] = e.dataset.k === 'monto' ? (e.value ? num(e.value) : null) : e.value.trim()));
    cont.querySelectorAll('[data-f]').forEach((e) => (o.feriados[e.dataset.f][e.dataset.k] = e.value.trim()));
    cont.querySelectorAll('[data-firma]').forEach((e) => (o.firmas[e.dataset.firma] = e.value.trim().toUpperCase()));
    cont.querySelectorAll('[data-par]').forEach((e) => {
      const k = e.dataset.par;
      o.parametros[k] = k === 'umbral' ? num(e.value) / 100 : k === 'jornada' ? num(e.value) : e.value;
    });
    o.feriados.sort((a, b) => (a.fecha < b.fecha ? -1 : 1));
  };
  const guardar = (msg = true) => {
    try { leerTodo(); } catch (e) { aviso(e.message, 'error'); return false; }
    cambio({ guardarYa: true });
    if (msg) aviso('Datos guardados.');
    return true;
  };
  cont.querySelector('#b-guardar').onclick = () => { if (guardar()) refrescar(); };
  const agregar = (lista, obj) => () => { if (!guardar(false)) return; o[lista].push(obj); cambio(); refrescar(); };
  cont.querySelector('#b-resp').onclick = agregar('responsables', { cargo: '', nombre: '', profesion: '' });
  cont.querySelector('#b-ind').onclick = agregar('indirectos', { concepto: 'NUEVO COSTO INDIRECTO', monto: 0 });
  cont.querySelector('#b-doc').onclick = agregar('documentos', { documento: '', resolucion: '', fecha: '', monto: null });
  cont.querySelector('#b-fer').onclick = agregar('feriados', { fecha: '', descripcion: '' });
  const borrar = (attr, lista) => cont.querySelectorAll(`[${attr}]`).forEach((b) => (b.onclick = async () => {
    if (!guardar(false)) return;
    if (!(await confirmar('¿Eliminar esta fila?'))) return;
    o[lista].splice(Number(b.getAttribute(attr)), 1);
    cambio();
    refrescar();
  }));
  borrar('data-rdel', 'responsables'); borrar('data-idel', 'indirectos'); borrar('data-ddel', 'documentos'); borrar('data-fdel', 'feriados');
  cont.querySelector('#b-pct').onclick = () => {
    if (!guardar(false)) return;
    if (!cd) { aviso('Primero cargue el presupuesto (C.D.).', 'alerta'); return; }
    const txt = prompt('Escriba los porcentajes del C.D. de cada indirecto, en orden, separados por ; (ej. 10;5;2):', o.indirectos.map((x) => fmt((x.monto / cd) * 100, 2)).join(';'));
    if (txt === null) return;
    txt.split(';').forEach((p, i) => { if (o.indirectos[i] && p.trim() !== '') o.indirectos[i].monto = Math.round(cd * num(p)) / 100; });
    cambio();
    refrescar();
  };
  cont.querySelectorAll('input,select').forEach((e) => e.addEventListener('change', () => guardar(false)));
}
