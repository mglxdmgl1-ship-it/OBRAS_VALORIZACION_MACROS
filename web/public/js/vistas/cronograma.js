// Cronograma valorizado (4.13): programado por partida y mes; generación automática e importación.
import { S, cambio } from '../store.js';
import { esc, fmt, num, normal, nombreMes, mesSiguiente, mesDe, hoy, r2 } from '../util.js';
import { icono, confirmar, aviso, elegirArchivo, trabajando } from '../ui.js';
import { cronogramaLineal } from '../calc.js';
import { leerLibro, importarCronograma } from '../importar.js';
import { rCronograma, rPlantilla } from '../reportes.js';
import { cab, btnExport, enlazarExport, sinPresupuesto } from './comun.js';
import { refrescar } from '../app.js';

let q = '', sub = '', modo = 'monto';

export function render(cont) {
  const o = S.obra, m = S.motor;
  if (!m.partidasP.length) { cont.innerHTML = cab('Cronograma valorizado', '') + sinPresupuesto(); return; }
  const meses = m.cronoMeses();
  const nq = normal(q);
  let filas = o.partidas;
  if (sub) filas = filas.filter((p) => m.raiz(p.item) === sub);
  if (nq) { const keep = new Set(); filas.forEach((p) => { if (p.tipo === 'P' && normal(p.item + ' ' + p.descripcion).includes(nq)) { keep.add(p.item); let pa = m.padre.get(p.item); while (pa) { keep.add(pa); pa = m.padre.get(pa); } } }); filas = filas.filter((p) => keep.has(p.item)); }
  const porMes = meses.map((ym) => m.sumarArbol((p) => m.programadoItemMes(p.item, ym)));
  const par = m.parciales();
  const totMes = meses.map((ym) => m.programadoMes(ym));
  const totGen = r2(totMes.reduce((a, b) => a + b, 0));
  const cd = m.costoDirecto();
  const difs = m.partidasP.filter((p) => meses.length && Math.abs(meses.reduce((s, ym) => s + m.programadoItemMes(p.item, ym), 0) - m.parcial(p)) > 0.05).length;
  let acum = 0;
  cont.innerHTML = `
    ${cab('Cronograma valorizado (4.13)', `Monto programado por partida y mes. Se usa para el cuadro 4.5 (programado vs ejecutado) y la curva S. Total programado S/ ${fmt(totGen)} de un C.D. de S/ ${fmt(cd)}.`,
      `<button class="btn" id="b-lin">${icono('cronograma')}Generar automático</button><button class="btn btn-oro" id="b-imp">${icono('importar')}Importar Excel</button><button class="btn" id="b-plant">${icono('excel')}Plantilla</button>${btnExport('cr', '4.13')}`)}
    ${!meses.length ? `<div class="panel"><div class="vacio"><h3>Sin cronograma</h3><p>Genérelo automáticamente (reparte cada partida en los meses del plazo según los días) e incorpore luego los ajustes, o impórtelo desde Excel (columna ITEM y una columna por mes).</p></div></div>` : `
    <div class="tarjetas">${meses.map((ym, i) => { acum += totMes[i]; return `<div class="tarjeta"><div class="et">${nombreMes(ym)}</div><div class="val" style="font-size:18px">S/ ${fmt(totMes[i])}</div><div class="sub">Acumulado ${fmt(cd ? (acum / cd) * 100 : 0)} %</div></div>`; }).join('')}</div>
    ${difs ? `<div class="nota nota-alerta">${difs} partida(s) tienen un programado total distinto a su parcial del presupuesto.</div>` : ''}
    <div class="filtros"><input type="search" class="buscar" id="q" placeholder="Buscar partida…" value="${esc(q)}">
      <select id="f-sub"><option value="">Todos los subpresupuestos</option>${m.raices.map((it) => `<option value="${esc(it)}" ${sub === it ? 'selected' : ''}>${esc(it)} ${esc(m.porItem.get(it).descripcion.slice(0, 40))}</option>`).join('')}</select>
      <div class="grupo-btn"><button data-modo="monto" class="${modo === 'monto' ? 'activo' : ''}">S/</button><button data-modo="pct" class="${modo === 'pct' ? 'activo' : ''}">% de la partida</button></div>
      <span style="flex:1"></span><button class="btn btn-sm" id="b-mas">${icono('mas')}Agregar mes</button><button class="btn btn-sm btn-peligro" id="b-menos">Quitar último mes</button><button class="btn btn-sm btn-peligro" id="b-borrar">${icono('borrar')}Borrar</button></div>
    <div class="tabla-cont"><table class="t" id="t-cr"><thead><tr><th>Ítem</th><th>Descripción</th><th class="n">Parcial S/</th>${meses.map((ym) => `<th class="n">${nombreMes(ym, true)}</th>`).join('')}<th class="n">Total</th><th class="n">Diferencia</th></tr></thead><tbody>
    ${filas.map((p) => {
      const niv = m.nivel(p.item);
      if (p.tipo === 'T') return `<tr class="tit ${niv === 1 ? 'tit-1' : ''}"><td class="item">${esc(p.item)}</td><td class="desc"><span class="sangria-${Math.min(niv, 7)}">${esc(p.descripcion)}</span></td><td class="n">${fmt(par.get(p.item))}</td>${meses.map((ym, i) => `<td class="n">${fmt(porMes[i].get(p.item), 2, true)}</td>`).join('')}<td class="n">${fmt(porMes.reduce((s, x) => s + (x.get(p.item) || 0), 0))}</td><td></td></tr>`;
      const pp = m.parcial(p);
      const vals = meses.map((ym) => m.programadoItemMes(p.item, ym));
      const tot = r2(vals.reduce((a, b) => a + b, 0));
      return `<tr data-it="${esc(p.item)}"><td class="item">${esc(p.item)}</td><td class="desc"><span class="sangria-${Math.min(niv, 7)}">${esc(p.descripcion)}</span></td><td class="n">${fmt(pp)}</td>${vals.map((v, i) => `<td><input class="n" data-m="${i}" value="${v ? (modo === 'pct' ? fmt(pp ? (v / pp) * 100 : 0, 2) : fmt(v)) : ''}" style="width:92px"></td>`).join('')}<td class="n">${fmt(tot)}</td><td class="n ${Math.abs(tot - pp) > 0.05 ? 'neg' : 'tenue'}">${fmt(r2(tot - pp), 2, true)}</td></tr>`;
    }).join('')}
    </tbody><tfoot><tr class="tot"><td></td><td>COSTO DIRECTO</td><td class="n">${fmt(cd)}</td>${totMes.map((v) => `<td class="n">${fmt(v)}</td>`).join('')}<td class="n">${fmt(totGen)}</td><td class="n">${fmt(r2(totGen - cd))}</td></tr></tfoot></table></div>`}`;
  enlazarExport(cont, { cr: () => rCronograma() });
  cont.querySelector('#b-plant').onclick = async () => { const { aExcel } = await import('../exportar.js'); aExcel(rPlantilla('cronograma')); };
  cont.querySelector('#b-lin').onclick = async () => {
    if (meses.length && !(await confirmar('¿Reemplazar el cronograma actual por uno generado automáticamente (proporcional a los días de cada mes del plazo)?', { peligro: true, si: 'Reemplazar' }))) return;
    const c = cronogramaLineal(o, m);
    if (!c) { aviso('Complete la fecha de inicio y el plazo en Datos de la obra.', 'alerta'); return; }
    o.cronograma = c; cambio(); aviso(`Cronograma generado: ${c.meses.length} meses.`); refrescar();
  };
  cont.querySelector('#b-imp').onclick = async () => {
    const f = await elegirArchivo('.xlsx,.xlsm,.csv');
    if (!f) return;
    try {
      const h = await trabajando('Leyendo…', () => leerLibro(f));
      let r = null, err = null;
      for (const hoja of h) { try { r = importarCronograma(o, hoja.filas); break; } catch (e) { err = e; } }
      if (!r) throw err;
      cambio({ guardarYa: true }); aviso(`${r.n} partida(s) y ${r.meses} mes(es) importados.`); refrescar();
    } catch (e) { aviso(e.message, 'error', 8000); }
  };
  if (!meses.length) return;
  let t;
  cont.querySelector('#q').oninput = (e) => { clearTimeout(t); t = setTimeout(() => { q = e.target.value; refrescar().then(() => { const el = document.querySelector('#q'); el?.focus(); el?.setSelectionRange(el.value.length, el.value.length); }); }, 300); };
  cont.querySelector('#f-sub').onchange = (e) => { sub = e.target.value; refrescar(); };
  cont.querySelectorAll('[data-modo]').forEach((b) => (b.onclick = () => { modo = b.dataset.modo; refrescar(); }));
  cont.querySelector('#b-mas').onclick = () => { const ult = meses[meses.length - 1] || mesDe(o.datos.fecha_inicio || hoy()); o.cronograma.meses.push(mesSiguiente(ult)); Object.values(o.cronograma.montos).forEach((a) => a.push(0)); cambio(); refrescar(); };
  cont.querySelector('#b-menos').onclick = async () => { if (!(await confirmar('¿Quitar el último mes y sus montos?', { peligro: true }))) return; o.cronograma.meses.pop(); Object.values(o.cronograma.montos).forEach((a) => a.pop()); cambio(); refrescar(); };
  cont.querySelector('#b-borrar').onclick = async () => { if (!(await confirmar('¿Borrar todo el cronograma?', { peligro: true }))) return; o.cronograma = { meses: [], montos: {} }; cambio(); refrescar(); };
  cont.querySelector('#t-cr tbody').addEventListener('change', (e) => {
    const el = e.target.closest('[data-m]');
    if (!el) return;
    const it = el.closest('tr').dataset.it;
    const p = m.porItem.get(it);
    const arr = o.cronograma.montos[it] || (o.cronograma.montos[it] = meses.map(() => 0));
    const v = num(el.value);
    arr[Number(el.dataset.m)] = modo === 'pct' ? r2((m.parcial(p) * v) / 100) : r2(v);
    cambio(); refrescar();
  });
}
