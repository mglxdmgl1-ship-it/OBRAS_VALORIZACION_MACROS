// Presupuesto: árbol de títulos y partidas, edición, importación desde Excel (S10, Delphin…).
import { S, cambio } from '../store.js';
import { esc, fmt, num, normal, cmpItem, normItem, r2 } from '../util.js';
import { icono, ventana, formulario, confirmar, aviso, elegirArchivo, trabajando } from '../ui.js';
import { leerLibro, analizarPresupuesto } from '../importar.js';
import { rPresupuesto, rPlantilla } from '../reportes.js';
import { cab, btnExport, enlazarExport } from './comun.js';
import { refrescar, ir } from '../app.js';

let filtro = { q: '', sub: '', soloDif: false };

export function render(cont) {
  const o = S.obra, m = S.motor;
  const par = m.parciales();
  const nq = normal(filtro.q);
  let visibles = o.partidas;
  if (filtro.sub) visibles = visibles.filter((p) => m.raiz(p.item) === filtro.sub);
  if (nq || filtro.soloDif) {
    const keep = new Set();
    for (const p of visibles) {
      const okQ = !nq || normal(p.item + ' ' + p.descripcion).includes(nq);
      const okD = !filtro.soloDif || (p.tipo === 'P' && m.acu(p.item).length && Math.abs(m.puACU(p.item) - (Number(p.pu) || 0)) >= 0.01);
      if (okQ && okD) { keep.add(p.item); let pa = m.padre.get(p.item); while (pa) { keep.add(pa); pa = m.padre.get(pa); } }
    }
    visibles = visibles.filter((p) => keep.has(p.item));
  }
  const nP = m.partidasP.length;
  cont.innerHTML = `
    ${cab('Presupuesto', `${nP} partidas y ${o.partidas.length - nP} títulos. Costo directo <b>S/ ${fmt(m.costoDirecto())}</b> · presupuesto total <b>S/ ${fmt(m.presupuestoTotal())}</b>. Puede digitarlo o importarlo desde Excel.`,
      `<button class="btn btn-oro" id="b-imp">${icono('importar')}Importar Excel</button><button class="btn" id="b-tit">${icono('mas')}Título</button><button class="btn btn-primario" id="b-par">${icono('mas')}Partida</button>${btnExport('pres')}`)}
    ${!o.partidas.length ? `<div class="panel"><div class="vacio"><h3>Presupuesto vacío</h3><p>Importe el presupuesto exportado de S10, Delphin, Kennedy u otro programa (Excel .xlsx con columnas ITEM, DESCRIPCIÓN, UND, METRADO, PRECIO, PARCIAL), o agregue títulos y partidas a mano.</p><p><button class="btn btn-oro" id="b-imp2">${icono('importar')}Importar presupuesto</button> <button class="btn" id="b-plant">${icono('excel')}Descargar plantilla</button></p></div></div>` : `
    <div class="filtros">
      <input type="search" class="buscar" id="f-q" placeholder="Buscar ítem o descripción…" value="${esc(filtro.q)}">
      <select id="f-sub"><option value="">Todos los subpresupuestos</option>${m.raices.map((it) => `<option value="${esc(it)}" ${filtro.sub === it ? 'selected' : ''}>${esc(it)} ${esc(m.porItem.get(it).descripcion.slice(0, 50))}</option>`).join('')}</select>
      <label class="chk"><input type="checkbox" id="f-dif" ${filtro.soloDif ? 'checked' : ''}> Solo P.U. distinto al A.C.U.</label>
      <span class="tenue" style="font-size:12px">${visibles.length} fila(s)</span>
      <span style="flex:1"></span>
      <button class="btn btn-sm" id="b-usarpu">Usar P.U. del A.C.U.…</button>
      <button class="btn btn-sm" id="b-plant">${icono('excel')}Plantilla</button>
      <button class="btn btn-sm btn-peligro" id="b-vaciar">${icono('borrar')}Vaciar</button>
    </div>
    <div class="tabla-cont"><table class="t" id="t-pres"><thead><tr><th>Ítem</th><th>Descripción</th><th>Und</th><th class="n">Metrado</th><th class="n">P.U. S/</th><th class="n">Parcial S/</th><th class="n" title="Suma del A.C.U.">P.U. A.C.U.</th><th class="n">Rend.</th><th></th></tr></thead><tbody>
    ${visibles.map((p) => fila(p, par, m)).join('')}
    </tbody><tfoot><tr class="tot"><td></td><td>COSTO DIRECTO</td><td></td><td></td><td></td><td class="n" id="tot-cd">${fmt(m.costoDirecto())}</td><td colspan="3"></td></tr></tfoot></table></div>`}`;

  enlazarExport(cont, { pres: () => rPresupuesto() });
  cont.querySelector('#b-imp').onclick = importar;
  cont.querySelector('#b-imp2')?.addEventListener('click', importar);
  cont.querySelectorAll('#b-plant').forEach((b) => (b.onclick = async () => { const { aExcel } = await import('../exportar.js'); aExcel(rPlantilla('presupuesto')); }));
  cont.querySelector('#b-tit').onclick = () => editar(null, 'T');
  cont.querySelector('#b-par').onclick = () => editar(null, 'P');
  const q = cont.querySelector('#f-q');
  if (q) {
    let t;
    q.oninput = () => { clearTimeout(t); t = setTimeout(() => { filtro.q = q.value; refrescar().then(() => { const e = document.querySelector('#f-q'); e.focus(); e.setSelectionRange(e.value.length, e.value.length); }); }, 250); };
    cont.querySelector('#f-sub').onchange = (e) => { filtro.sub = e.target.value; refrescar(); };
    cont.querySelector('#f-dif').onchange = (e) => { filtro.soloDif = e.target.checked; refrescar(); };
    cont.querySelector('#b-vaciar').onclick = async () => {
      if (!(await confirmar('¿Borrar <b>todo el presupuesto</b> (títulos, partidas, A.C.U., cronograma y metrados)? Haga antes un respaldo.', { peligro: true, si: 'Borrar todo' }))) return;
      Object.assign(o, { partidas: [], acu: {}, metrados: {}, detalles: {}, cronograma: { meses: [], montos: {} }, anterior_inicial: {} });
      o.periodos.forEach((p) => { p.manual = {}; p.cerrado = p.cerrado ? {} : null; });
      cambio(); refrescar();
    };
    cont.querySelector('#b-usarpu').onclick = usarPUdelACU;
  }
  const tb = cont.querySelector('#t-pres tbody');
  if (!tb) return;
  tb.addEventListener('change', (e) => {
    const inp = e.target.closest('input[data-k]');
    if (!inp) return;
    const p = m.porItem.get(inp.closest('tr').dataset.it);
    p[inp.dataset.k] = num(inp.value);
    cambio();
    const par2 = S.motor.parciales();
    inp.closest('tr').querySelector('.parcial').textContent = fmt(par2.get(p.item));
    inp.value = fmt(p[inp.dataset.k], inp.dataset.k === 'metrado' ? 2 : 2);
    // actualizar títulos visibles y total
    tb.querySelectorAll('tr.tit').forEach((tr) => (tr.querySelector('.parcial').textContent = fmt(par2.get(tr.dataset.it))));
    cont.querySelector('#tot-cd').textContent = fmt(S.motor.costoDirecto());
  });
  tb.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-acc]');
    if (!b) return;
    const it = b.closest('tr').dataset.it;
    if (b.dataset.acc === 'ed') editar(it);
    if (b.dataset.acc === 'acu') ir('acu', { item: it });
    if (b.dataset.acc === 'hija') editar(null, 'P', it);
    if (b.dataset.acc === 'del') eliminar(it);
  });
}

function fila(p, par, m) {
  const niv = m.nivel(p.item);
  if (p.tipo === 'T') {
    return `<tr class="tit ${niv === 1 ? 'tit-1' : ''}" data-it="${esc(p.item)}"><td class="item">${esc(p.item)}</td><td class="desc"><span class="sangria-${Math.min(niv, 7)}">${esc(p.descripcion)}</span></td><td></td><td></td><td></td><td class="n parcial">${fmt(par.get(p.item))}</td><td></td><td></td>
      <td class="acc"><button class="btn-icono" data-acc="hija" title="Agregar partida dentro">${icono('mas')}</button><button class="btn-icono" data-acc="ed" title="Editar">${icono('editar')}</button><button class="btn-icono peligro" data-acc="del" title="Eliminar">${icono('borrar')}</button></td></tr>`;
  }
  const nA = m.acu(p.item).length;
  const puA = nA ? m.puACU(p.item) : null;
  const dif = puA !== null && Math.abs(puA - (Number(p.pu) || 0)) >= 0.01;
  return `<tr data-it="${esc(p.item)}"><td class="item">${esc(p.item)}</td><td class="desc"><span class="sangria-${Math.min(niv, 7)}">${esc(p.descripcion)}</span>${p.origen && p.origen !== 'ET' ? ` <span class="badge badge-alerta">${esc(p.origen.replace('_', ' '))}</span>` : ''}</td><td>${esc(p.und)}</td>
    <td><input class="n" data-k="metrado" value="${fmt(p.metrado)}"></td><td><input class="n" data-k="pu" value="${fmt(p.pu)}"></td><td class="n parcial">${fmt(par.get(p.item))}</td>
    <td class="n ${dif ? 'neg' : 'tenue'}" title="${nA} insumo(s)">${puA !== null ? fmt(puA) : '<span class="badge badge-alerta">sin A.C.U.</span>'}</td><td class="n tenue">${p.rendimiento ?? ''}</td>
    <td class="acc"><button class="btn-icono" data-acc="acu" title="Ver A.C.U.">${icono('acu')}</button><button class="btn-icono" data-acc="ed" title="Editar">${icono('editar')}</button><button class="btn-icono peligro" data-acc="del" title="Eliminar">${icono('borrar')}</button></td></tr>`;
}

/** Cambia el código de un ítem en todas las tablas de la obra. */
export function renombrarItem(o, viejo, nuevo) {
  if (viejo === nuevo) return;
  const mover = (obj) => { if (obj && Object.prototype.hasOwnProperty.call(obj, viejo)) { obj[nuevo] = obj[viejo]; delete obj[viejo]; } };
  mover(o.acu); mover(o.cronograma.montos); mover(o.anterior_inicial);
  Object.values(o.metrados).forEach(mover);
  o.periodos.forEach((p) => { mover(p.manual); mover(p.cerrado); });
  o.movimientos.forEach((x) => { if (x.partida === viejo) x.partida = nuevo; if (x.partida_cubre === viejo) x.partida_cubre = nuevo; });
  o.partes.forEach((x) => { if (x.partida === viejo) x.partida = nuevo; });
  for (const k of Object.keys(o.detalles)) { const [f, it] = k.split('|'); if (it === viejo) { o.detalles[f + '|' + nuevo] = o.detalles[k]; delete o.detalles[k]; } }
}
function ordenar(o) {
  o.partidas.sort((a, b) => cmpItem(a.item, b.item));
}
function siguienteItem(o, padre) {
  const hijos = o.partidas.filter((p) => (padre ? p.item.startsWith(padre + '.') && p.item.split('.').length === padre.split('.').length + 1 : !p.item.includes('.')));
  let n = 0;
  hijos.forEach((p) => { n = Math.max(n, parseInt(p.item.split('.').pop(), 10) || 0); });
  return (padre ? padre + '.' : '') + String(n + 1).padStart(2, '0');
}

function editar(item, tipo = 'P', padre = null) {
  const o = S.obra;
  const p = item ? S.motor.porItem.get(item) : null;
  const t = p ? p.tipo : tipo;
  const val = p ? { ...p } : { item: siguienteItem(o, padre), tipo: t, origen: 'ET', cuadrillas: 1 };
  const campos = [
    { id: 'item', et: 'Ítem', req: true, ph: '01.02.03' },
    { id: 'tipo', et: 'Tipo', tipo: 'select', ops: [['T', 'Título'], ['P', 'Partida']] },
    { id: 'origen', et: 'Origen', tipo: 'select', ops: [['ET', 'Expediente técnico'], ['MAYOR_METRADO', 'Mayor metrado'], ['PARTIDA_NUEVA', 'Partida nueva'], ['DEDUCTIVO', 'Deductivo']] },
    { id: 'descripcion', et: 'Descripción', req: true, col: 3 },
    { id: 'und', et: 'Unidad', ph: 'm2, m3, kg, und, glb…' },
    { id: 'metrado', et: 'Metrado', tipo: 'number' },
    { id: 'pu', et: 'Precio unitario (S/)', tipo: 'number' },
    { id: 'rendimiento', et: 'Rendimiento (und/día)', tipo: 'number', ayuda: 'Para la alerta «supera rendimiento» y la cuadrilla' },
    { id: 'cuadrillas', et: 'N° de cuadrillas', tipo: 'number', ayuda: 'Máximo diario = rendimiento × cuadrillas' },
  ];
  formulario({
    titulo: p ? `Editar ${t === 'T' ? 'título' : 'partida'} ${item}` : `Nuevo ${t === 'T' ? 'título' : 'partida'}`, campos, val,
    alGuardar: (v) => {
      v.item = normItem(v.item);
      if (!/^\d{1,4}(\.\d{1,4})*$/.test(v.item)) throw new Error('El ítem debe tener el formato 01.02.03');
      if (v.item !== item && S.motor.porItem.has(v.item)) throw new Error('Ya existe el ítem ' + v.item);
      v.descripcion = v.descripcion.toUpperCase() === v.descripcion ? v.descripcion : v.descripcion;
      if (v.tipo === 'T') { v.und = ''; v.metrado = null; v.pu = null; v.rendimiento = null; }
      else if (!v.und) throw new Error('Una partida necesita unidad.');
      if (p) {
        renombrarItem(o, item, v.item);
        Object.assign(p, v);
      } else {
        o.partidas.push({ ...v, metrado: v.metrado ?? 0, pu: v.pu ?? 0 });
        // títulos padre que falten
        const seg = v.item.split('.');
        for (let n = 1; n < seg.length; n++) {
          const c = seg.slice(0, n).join('.');
          if (!o.partidas.some((x) => x.item === c)) o.partidas.push({ item: c, tipo: 'T', descripcion: 'TÍTULO ' + c, und: '', metrado: null, pu: null, origen: 'ET' });
        }
      }
      ordenar(o);
      cambio();
      aviso('Guardado.');
      refrescar();
    },
  });
}

async function eliminar(item) {
  const o = S.obra;
  const p = S.motor.porItem.get(item);
  const desc = o.partidas.filter((x) => x.item === item || x.item.startsWith(item + '.'));
  const conMet = desc.filter((x) => Object.values(o.metrados).some((d) => Number(d[x.item]) > 0)).length;
  if (!(await confirmar(`¿Eliminar ${p.tipo === 'T' ? `el título <b>${esc(item)}</b> y sus ${desc.length - 1} ítem(s)` : `la partida <b>${esc(item)}</b>`}?${conMet ? `<br><br><b class="neg">${conMet} partida(s) tienen metrados diarios registrados</b>; también se borrarán.` : ''}`, { peligro: true, si: 'Eliminar' }))) return;
  const set = new Set(desc.map((x) => x.item));
  o.partidas = o.partidas.filter((x) => !set.has(x.item));
  for (const it of set) {
    delete o.acu[it]; delete o.cronograma.montos[it]; delete o.anterior_inicial[it];
    Object.values(o.metrados).forEach((d) => delete d[it]);
    o.periodos.forEach((pp) => { delete pp.manual?.[it]; if (pp.cerrado) delete pp.cerrado[it]; });
  }
  cambio();
  aviso('Eliminado.');
  refrescar();
}

function usarPUdelACU() {
  const m = S.motor;
  const difs = m.partidasP.filter((p) => m.acu(p.item).length && Math.abs(m.puACU(p.item) - (Number(p.pu) || 0)) >= 0.01);
  if (!difs.length) { aviso('Todas las partidas con A.C.U. tienen el mismo P.U. que el análisis.'); return; }
  ventana({
    titulo: 'Usar el P.U. calculado en el A.C.U.', ancho: '820px',
    cuerpo: `<p>${difs.length} partida(s) tienen un P.U. distinto a la suma de su A.C.U. Marque las que quiere actualizar (cambia el costo directo).</p>
      <div class="tabla-cont baja"><table class="t"><thead><tr><th><input type="checkbox" id="chk-all"></th><th>Ítem</th><th>Partida</th><th class="n">P.U. actual</th><th class="n">P.U. A.C.U.</th><th class="n">Diferencia parcial</th></tr></thead><tbody>
      ${difs.map((p) => `<tr><td><input type="checkbox" data-it="${esc(p.item)}"></td><td class="item">${esc(p.item)}</td><td>${esc(p.descripcion)}</td><td class="n">${fmt(p.pu)}</td><td class="n">${fmt(m.puACU(p.item))}</td><td class="n">${fmt(r2((m.puACU(p.item) - p.pu) * p.metrado))}</td></tr>`).join('')}
      </tbody></table></div>`,
    botones: [{ txt: 'Cancelar' }, { txt: 'Actualizar marcadas', clase: 'btn-primario', accion: (c, raiz) => {
      const its = [...raiz.querySelectorAll('input[data-it]:checked')].map((e) => e.dataset.it);
      its.forEach((it) => { const p = m.porItem.get(it); p.pu = m.puACU(it); });
      cambio(); aviso(`${its.length} P.U. actualizado(s).`); refrescar();
    } }],
    alAbrir: (raiz) => raiz.querySelector('#chk-all').addEventListener('change', (e) => raiz.querySelectorAll('input[data-it]').forEach((c) => (c.checked = e.target.checked))),
  });
}

export async function importar() {
  const f = await elegirArchivo('.xlsx,.xlsm,.csv');
  if (!f) return;
  let res;
  try {
    res = await trabajando('Leyendo ' + f.name + '…', async () => analizarPresupuesto(await leerLibro(f)));
  } catch (e) { aviso(e.message, 'error', 8000); return; }
  const o = S.obra;
  const nP = res.partidas.filter((p) => p.tipo === 'P').length;
  const cd = r2(res.partidas.filter((p) => p.tipo === 'P').reduce((s, p) => s + r2(p.metrado * p.pu), 0));
  ventana({
    titulo: 'Vista previa del presupuesto: ' + f.name, ancho: '980px',
    cuerpo: `<div class="tarjetas"><div class="tarjeta"><div class="et">Hoja</div><div class="val" style="font-size:16px">${esc(res.hoja)}</div></div><div class="tarjeta"><div class="et">Partidas</div><div class="val">${nP}</div></div><div class="tarjeta"><div class="et">Títulos</div><div class="val">${res.partidas.length - nP}</div></div><div class="tarjeta"><div class="et">Costo directo</div><div class="val">S/ ${fmt(cd)}</div></div></div>
      ${res.avisos.length ? `<div class="nota nota-alerta"><b>Observaciones (${res.avisos.length}):</b><ul class="lista-simple">${res.avisos.slice(0, 15).map((a) => `<li>${esc(a)}</li>`).join('')}${res.avisos.length > 15 ? `<li>… y ${res.avisos.length - 15} más</li>` : ''}</ul></div>` : '<div class="nota nota-ok">El archivo se leyó sin observaciones.</div>'}
      <div class="tabla-cont baja"><table class="t"><thead><tr><th>Ítem</th><th>Descripción</th><th>Und</th><th class="n">Metrado</th><th class="n">P.U.</th><th class="n">Parcial</th></tr></thead><tbody>
      ${res.partidas.slice(0, 120).map((p) => `<tr class="${p.tipo === 'T' ? 'tit' : ''}"><td class="item">${esc(p.item)}</td><td>${esc(p.descripcion)}</td><td>${esc(p.und)}</td><td class="n">${p.tipo === 'P' ? fmt(p.metrado) : ''}</td><td class="n">${p.tipo === 'P' ? fmt(p.pu) : ''}</td><td class="n">${p.tipo === 'P' ? fmt(r2(p.metrado * p.pu)) : ''}</td></tr>`).join('')}
      ${res.partidas.length > 120 ? `<tr><td colspan="6" class="tenue">… ${res.partidas.length - 120} filas más</td></tr>` : ''}</tbody></table></div>
      <div class="form-grid" style="margin-top:12px">
        <label class="campo c2"><span>¿Qué hacer con el presupuesto actual (${o.partidas.length} filas)?</span><select id="imp-modo"><option value="reemplazar">Reemplazar el presupuesto (conserva A.C.U. y metrados de los ítems que coincidan)</option><option value="fusionar">Fusionar: actualizar los ítems que existan y agregar los nuevos</option></select></label>
        ${res.indirectos.length ? `<label class="campo"><span>Resumen encontrado al final</span><label class="chk"><input type="checkbox" id="imp-ind"> Usar como costos indirectos (${res.indirectos.length})</label><small>${esc(res.indirectos.map((x) => x.concepto + ': ' + fmt(x.monto)).join(' · ').slice(0, 240))}</small></label>` : ''}
      </div>`,
    botones: [{ txt: 'Cancelar' }, { txt: 'Importar presupuesto', clase: 'btn-primario', icono: 'importar', accion: (c, raiz) => {
      const modo = raiz.querySelector('#imp-modo').value;
      if (modo === 'reemplazar') {
        const nuevos = new Set(res.partidas.map((p) => p.item));
        const viejos = new Map(o.partidas.map((p) => [p.item, p]));
        res.partidas.forEach((p) => { const v = viejos.get(p.item); if (v && v.rendimiento && !p.rendimiento) p.rendimiento = v.rendimiento; });
        o.partidas = res.partidas;
        for (const k of Object.keys(o.acu)) if (!nuevos.has(k)) delete o.acu[k];
      } else {
        const idx = new Map(o.partidas.map((p) => [p.item, p]));
        for (const p of res.partidas) { const v = idx.get(p.item); if (v) Object.assign(v, { ...p, rendimiento: p.rendimiento ?? v.rendimiento }); else o.partidas.push(p); }
        ordenar(o);
      }
      const usarInd = raiz.querySelector('#imp-ind')?.checked;
      if (usarInd) o.indirectos = res.indirectos.filter((x) => !/COSTO DIRECTO|TOTAL|I\.?G\.?V/i.test(x.concepto)).map((x) => ({ concepto: x.concepto, monto: x.monto }));
      cambio({ guardarYa: true });
      aviso(`Presupuesto importado: ${nP} partidas.`);
      refrescar();
    } }],
  });
}
