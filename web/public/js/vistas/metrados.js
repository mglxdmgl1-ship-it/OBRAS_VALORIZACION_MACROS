// Control diario de metrados: planilla del mes (RESUMEN), planilla del día con detalle, insumos del mes.
import { S, cambio } from '../store.js';
import { esc, fmt, cant, num, normal, hoy, r2, diasDelMes, finDeMes, mesDe, nombreMes, diaSemana, mesesEntre, fFecha, sumarDias, DIAS } from '../util.js';
import { icono, ventana, aviso, elegirArchivo, trabajando, confirmar } from '../ui.js';
import { leerLibro, importarMetrados } from '../importar.js';
import { rPlanillaDia, rMetradosMes, rPlantilla } from '../reportes.js';
import { cab, btnExport, enlazarExport, chipEstado, leyendaEstados, elegirPartidas, sinPresupuesto, mesesDeObra } from './comun.js';
import { refrescar, ir } from '../app.js';

let tab = 'mes', mes = null, dia = null, q = '', sub = '', soloCon = true, verFuera = false;
const extra = new Set();
const extraDia = new Map();

function mesesPosibles() {
  const o = S.obra, m = S.motor;
  const set = new Set(mesesDeObra());
  if (m.inicio && m.termino) mesesEntre(m.inicio, m.termino).forEach((x) => set.add(x));
  set.add(mesDe(hoy()));
  return [...set].sort();
}

export function render(cont, params) {
  const m = S.motor;
  if (!m.partidasP.length) { cont.innerHTML = cab('Metrados diarios', '') + sinPresupuesto(); return; }
  if (params.dia) { dia = params.dia; tab = 'dia'; mes = mesDe(dia); }
  if (params.tab) tab = params.tab;
  if (!mes) {
    const fs = m.fechasConMetrado();
    mes = fs.length ? mesDe(fs[fs.length - 1]) : mesDe(m.inicio || hoy());
  }
  if (!dia) { const fs = m.fechasConMetrado().filter((f) => mesDe(f) === mes); dia = fs.length ? fs[fs.length - 1] : (hoy().startsWith(mes) ? hoy() : mes + '-01'); }
  cont.innerHTML = `
    ${cab('Metrados diarios', 'Digite el metrado ejecutado de cada partida por día. El programa calcula el consumo de insumos según el A.C.U. y marca cada partida-día según los ingresos al almacén / O/S, el rendimiento y el metrado del expediente (ModActualizar).',
      `<button class="btn" id="b-imp">${icono('importar')}Importar Excel</button><button class="btn" id="b-plant">${icono('excel')}Plantilla</button>`)}
    <div class="pestanas">
      <button data-tab="mes" class="${tab === 'mes' ? 'activo' : ''}">Planilla del mes</button>
      <button data-tab="dia" class="${tab === 'dia' ? 'activo' : ''}">Planilla del día</button>
      <button data-tab="insumos" class="${tab === 'insumos' ? 'activo' : ''}">Insumos del mes</button>
    </div><div id="cuerpo"></div>`;
  cont.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { tab = b.dataset.tab; history.replaceState(null, '', '#/metrados'); refrescar(); }));
  cont.querySelector('#b-plant').onclick = async () => { const { aExcel } = await import('../exportar.js'); aExcel(rPlantilla('metrados')); };
  cont.querySelector('#b-imp').onclick = importar;
  const cuerpo = cont.querySelector('#cuerpo');
  if (tab === 'mes') vistaMes(cuerpo);
  else if (tab === 'dia') vistaDia(cuerpo);
  else vistaInsumos(cuerpo);
}

// =================================================================== MES
function vistaMes(box) {
  const o = S.obra, m = S.motor;
  const nd = diasDelMes(mes);
  const ini = mes + '-01', fin = finDeMes(mes);
  const todas = Array.from({ length: nd }, (_, i) => `${mes}-${String(i + 1).padStart(2, '0')}`);
  const fueraP = (f) => (m.inicio && f < m.inicio) || (m.termino && f > m.termino);
  const nFuera = todas.filter(fueraP).length;
  const fechas = verFuera ? todas : todas.filter((f) => !fueraP(f) || m.partidasP.some((p) => m.metradoDia(f, p.item) > 0));
  const conMet = new Set(m.partidasP.filter((p) => m.sumaDiaria(p.item, ini, fin) > 0).map((p) => p.item));
  const nq = normal(q);
  let lista = m.partidasP.filter((p) => (!sub || m.raiz(p.item) === sub) && (!nq || normal(p.item + ' ' + p.descripcion).includes(nq)));
  if (soloCon && !nq) lista = lista.filter((p) => conMet.has(p.item) || extra.has(p.item));
  const limite = 250;
  const recortado = lista.length > limite;
  lista = lista.slice(0, limite);
  const fer = new Set((o.feriados || []).map((x) => x.fecha));
  const est = m.estadosMes(mes);
  const valDia = fechas.map((f) => r2(m.partidasP.reduce((s, p) => s + r2(m.metradoDia(f, p.item) * (Number(p.pu) || 0)), 0)));
  const clsDia = (f) => [diaSemana(f) === 0 ? 'dom' : '', fer.has(f) ? 'fer' : '', (m.inicio && f < m.inicio) || (m.termino && f > m.termino) ? 'fuera' : '', f === hoy() ? 'hoy' : ''].join(' ');
  box.innerHTML = `
    <div class="filtros">
      <select id="f-mes">${mesesPosibles().map((x) => `<option value="${x}" ${x === mes ? 'selected' : ''}>${nombreMes(x)}</option>`).join('')}</select>
      <input type="search" class="buscar" id="f-q" placeholder="Buscar partida (para agregar al mes)…" value="${esc(q)}">
      <select id="f-sub"><option value="">Todos los subpresupuestos</option>${m.raices.map((it) => `<option value="${esc(it)}" ${sub === it ? 'selected' : ''}>${esc(it)} ${esc(m.porItem.get(it).descripcion.slice(0, 40))}</option>`).join('')}</select>
      <label class="chk"><input type="checkbox" id="f-solo" ${soloCon ? 'checked' : ''}> Solo partidas con metrado</label>
      ${nFuera ? `<label class="chk"><input type="checkbox" id="f-fuera" ${verFuera ? 'checked' : ''}> Ver ${nFuera} día(s) fuera del plazo</label>` : ''}
      <button class="btn btn-sm" id="b-agregar">${icono('mas')}Agregar partidas</button>
      <span style="flex:1"></span>${btnExport('mes', '4.1')}
    </div>
    ${leyendaEstados()}
    ${recortado ? `<div class="nota nota-alerta">Se muestran las primeras ${limite} partidas. Use el buscador o el filtro de subpresupuesto.</div>` : ''}
    ${!lista.length ? `<div class="panel"><div class="vacio"><h3>Sin metrados en ${nombreMes(mes)}</h3><p>Use «Agregar partidas» o el buscador para elegir las partidas trabajadas, o vaya a la <a href="#/metrados?tab=dia">planilla del día</a>.</p></div></div>` : `
    <div class="tabla-cont"><table class="t grilla" id="t-mes"><thead><tr><th class="fijo">Ítem</th><th>Descripción</th><th>Und</th><th class="n">Met. E.T.</th><th class="n">Acum. ant.</th>
      ${fechas.map((f) => `<th class="n ${clsDia(f)}" style="cursor:pointer" data-ir="${f}" title="${DIAS[diaSemana(f)]} ${fFecha(f)}${fer.has(f) ? ' · feriado' : ''}">${f.slice(8)}<br><span style="font-weight:400;font-size:10px">${DIAS[diaSemana(f)].slice(0, 2)}</span></th>`).join('')}
      <th class="n">Mes</th><th class="n">Acum.</th><th class="n">%</th><th class="n">Saldo</th><th class="n">Valorizado S/</th><th>Estado del mes</th></tr></thead><tbody>
      ${lista.map((p) => filaMes(p, fechas, ini, fin, est, clsDia)).join('')}
    </tbody><tfoot><tr class="tot"><td class="fijo"></td><td>VALORIZADO DEL DÍA (C.D.)</td><td></td><td></td><td></td>${valDia.map((v) => `<td class="n" style="font-size:11px">${v ? fmt(v, 0) : ''}</td>`).join('')}<td></td><td></td><td></td><td></td><td class="n">${fmt(valDia.reduce((a, b) => a + b, 0))}</td><td></td></tr></tfoot></table></div>`}`;
  enlazarExport(box, { mes: () => rMetradosMes(mes) });
  box.querySelector('#f-mes').onchange = (e) => { mes = e.target.value; dia = null; extra.clear(); refrescar(); };
  let t;
  box.querySelector('#f-q').oninput = (e) => { clearTimeout(t); t = setTimeout(() => { q = e.target.value; refrescar().then(() => { const el = document.querySelector('#f-q'); el.focus(); el.setSelectionRange(el.value.length, el.value.length); }); }, 300); };
  box.querySelector('#f-sub').onchange = (e) => { sub = e.target.value; refrescar(); };
  box.querySelector('#f-solo').onchange = (e) => { soloCon = e.target.checked; refrescar(); };
  box.querySelector('#f-fuera')?.addEventListener('change', (e) => { verFuera = e.target.checked; refrescar(); });
  box.querySelector('#b-agregar').onclick = async () => { const its = await elegirPartidas({ titulo: 'Agregar partidas a la planilla del mes' }); its.forEach((x) => extra.add(x)); if (its.length) { soloCon = true; q = ''; refrescar(); } };
  box.querySelectorAll('[data-ir]').forEach((th) => (th.onclick = () => { dia = th.dataset.ir; tab = 'dia'; refrescar(); }));
  const tb = box.querySelector('#t-mes tbody');
  tb?.addEventListener('change', (e) => {
    const el = e.target.closest('input[data-f]');
    if (!el || !el.isConnected) return;
    const it = el.closest('tr').dataset.it, f = el.dataset.f;
    if (num(el.value) === S.motor.metradoDia(f, it)) return;
    setMetrado(f, it, num(el.value));
    const p = S.motor.porItem.get(it);
    const tr = el.closest('tr');
    const tmp = document.createElement('tbody');
    tmp.innerHTML = filaMes(p, fechas, ini, fin, S.motor.estadosMes(mes), clsDia);
    tr.replaceWith(tmp.firstElementChild);
    const sig = tb.querySelector(`tr[data-it="${CSS.escape(it)}"] input[data-f="${sumarDias(f, 1)}"]`);
    if (document.activeElement === document.body && sig) sig.focus();
  });
  tb?.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter' && e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    const el = e.target.closest('input[data-f]');
    if (!el) return;
    e.preventDefault();
    const tr = el.closest('tr');
    const dest = (e.key === 'ArrowUp' ? tr.previousElementSibling : tr.nextElementSibling)?.querySelector(`input[data-f="${el.dataset.f}"]`);
    el.dispatchEvent(new Event('change', { bubbles: true }));
    setTimeout(() => { const d2 = dest && tb.querySelector(`tr[data-it="${CSS.escape(dest.closest('tr').dataset.it)}"] input[data-f="${el.dataset.f}"]`); d2?.focus(); d2?.select(); }, 0);
  });
}
function filaMes(p, fechas, ini, fin, est, clsDia) {
  const m = S.motor;
  const ant = (Number(S.obra.anterior_inicial?.[p.item]) || 0) + m.sumaDiaria(p.item, null, sumarDias(ini, -1));
  const sm = m.sumaDiaria(p.item, ini, fin);
  const acu = ant + sm;
  const e = est.get(p.item);
  return `<tr data-it="${esc(p.item)}"><td class="item fijo">${esc(p.item)}</td><td><span class="desc-trunc" style="max-width:300px" title="${esc(p.descripcion)}">${esc(p.descripcion)}</span></td><td>${esc(p.und)}</td><td class="n">${fmt(p.metrado)}</td><td class="n tenue">${fmt(ant, 2, true)}</td>
    ${fechas.map((f) => { const v = m.metradoDia(f, p.item); const ed = v > 0 ? m.estadoDia(p.item, f) : null; return `<td class="dia ${clsDia(f)} ${ed ? ed.estado.clase : ''}" title="${ed ? esc(ed.estado.txt + '\n' + ed.motivos.join('\n')) : ''}"><input data-f="${f}" value="${v ? cant(v, 2) : ''}" inputmode="decimal"></td>`; }).join('')}
    <td class="n"><b>${fmt(sm, 2, true)}</b></td><td class="n">${fmt(acu)}</td><td class="n ${acu > p.metrado * 1.0001 ? 'neg' : ''}">${p.metrado ? fmt((acu / p.metrado) * 100, 1) + '%' : ''}</td><td class="n">${fmt(p.metrado - acu)}</td><td class="n">${fmt(r2(sm * (Number(p.pu) || 0)), 2, true)}</td><td>${e ? chipEstado(e, true) : ''}</td></tr>`;
}
function setMetrado(f, it, v) {
  const o = S.obra;
  if (!o.metrados[f]) o.metrados[f] = {};
  if (v > 0) o.metrados[f][it] = r2(v * 10000) / 10000;
  else { delete o.metrados[f][it]; delete o.detalles[f + '|' + it]; if (!Object.keys(o.metrados[f]).length) delete o.metrados[f]; }
  cambio();
}

// =================================================================== DÍA
function vistaDia(box) {
  const o = S.obra, m = S.motor;
  const items = [...new Set([...Object.entries(o.metrados[dia] || {}).filter(([, v]) => Number(v) > 0).map(([it]) => it), ...(extraDia.get(dia) || [])])].filter((it) => m.porItem.get(it));
  items.sort((a, b) => o.partidas.findIndex((p) => p.item === a) - o.partidas.findIndex((p) => p.item === b));
  const lineas = items.map((it) => ({ item: it, metrado: m.metradoDia(dia, it) })).filter((l) => l.metrado > 0);
  const consumo = m.consumo(lineas, { incluirMO: true });
  const tot = r2(items.reduce((s, it) => s + r2(m.metradoDia(dia, it) * (Number(m.porItem.get(it).pu) || 0)), 0));
  const nDia = m.inicio ? Math.round((Date.parse(dia) - Date.parse(m.inicio)) / 86400000) + 1 : null;
  const fuera = (m.inicio && dia < m.inicio) || (m.termino && dia > m.termino);
  const fer = (o.feriados || []).find((x) => x.fecha === dia);
  box.innerHTML = `
    <div class="filtros">
      <button class="btn btn-sm" id="d-ant" data-libre>‹</button><input type="date" id="d-fecha" value="${dia}"><button class="btn btn-sm" id="d-sig" data-libre>›</button>
      <span class="badge badge-azul">${DIAS[diaSemana(dia)]}</span>${nDia ? `<span class="badge">Día de obra N° ${nDia}</span>` : ''}${fer ? `<span class="badge badge-alerta">Feriado: ${esc(fer.descripcion)}</span>` : ''}${fuera ? '<span class="badge badge-error">Fuera del plazo</span>' : ''}
      <span style="flex:1"></span><button class="btn btn-primario btn-sm" id="d-add">${icono('mas')}Agregar partidas</button>${btnExport('dia', 'Planilla')}
    </div>
    <div class="panel"><div class="panel-cab"><h2>Planilla de metrados del ${fFecha(dia)}</h2><span class="badge badge-azul">Valorizado del día S/ ${fmt(tot)}</span></div>
    ${items.length ? `<div class="tabla-cont libre" style="border:0"><table class="t" id="t-dia"><thead><tr><th>Ítem</th><th>Descripción</th><th>Und</th><th class="n">Metrado del día</th><th></th><th class="n">Acum. al día</th><th class="n">Met. E.T.</th><th class="n">% acum.</th><th class="n">P.U.</th><th class="n">Parcial S/</th><th>Estado / alerta</th><th></th></tr></thead><tbody>
      ${items.map((it) => { const p = m.porItem.get(it); const md = m.metradoDia(dia, it); const e = m.estadoDia(it, dia) || { acumulado: m.acumuladoHasta(it, dia), motivos: ['Escriba el metrado del día'], estado: null }; const det = o.detalles[dia + '|' + it]; return `<tr data-it="${esc(it)}"><td class="item">${esc(it)}</td><td class="desc">${esc(p.descripcion)}</td><td>${esc(p.und)}</td>
        <td><input class="n" data-met value="${md ? cant(md, 4) : ''}" style="width:96px" ${det?.length ? 'readonly title="Calculado con el detalle"' : ''}></td><td><button class="btn btn-sm" data-det title="Planilla de metrados (elementos y medidas)">${icono('presupuesto')}${det?.length ? det.length : ''}</button></td>
        <td class="n">${fmt(e.acumulado)}</td><td class="n">${fmt(p.metrado)}</td><td class="n">${p.metrado ? fmt((e.acumulado / p.metrado) * 100, 1) + '%' : ''}</td><td class="n">${fmt(p.pu)}</td><td class="n">${fmt(r2(md * p.pu))}</td>
        <td>${e.estado ? chipEstado(e) : ''}${e.motivos.length ? `<div class="tenue" style="font-size:11px;max-width:380px">${esc(e.motivos.join(' · '))}</div>` : ''}</td><td><button class="btn-icono peligro" data-del title="Quitar">${icono('borrar')}</button></td></tr>`; }).join('')}
    </tbody></table></div>` : `<div class="vacio"><h3>Sin metrados este día</h3><p>Use «Agregar partidas» para registrar lo ejecutado.</p></div>`}</div>
    ${items.length ? `<div class="panel"><div class="panel-cab"><h2>Insumos del día según A.C.U. (metrado × cantidad)</h2><div class="acciones"><button class="btn btn-sm" id="d-calc">${icono('calc')}Abrir en la calculadora</button><button class="btn btn-sm btn-primario" id="d-sal">${icono('almacen')}Generar salidas del almacén</button><button class="btn btn-sm" id="d-par">${icono('equipos')}Generar partes de equipo</button></div></div>
      <div class="tabla-cont baja" style="border:0"><table class="t"><thead><tr><th>Insumo</th><th>Und</th><th>Tipo</th><th class="n">Cantidad</th><th class="n">Precio</th><th class="n">Parcial S/</th></tr></thead><tbody>
      ${consumo.map((x) => `<tr><td>${esc(x.insumo)}</td><td>${esc(x.und)}</td><td><span class="badge">${x.tipo}</span></td><td class="n">${cant(x.cantidad, 4)}</td><td class="n">${fmt(x.precio)}</td><td class="n">${fmt(x.monto)}</td></tr>`).join('')}
      </tbody></table></div></div>` : ''}`;
  enlazarExport(box, { dia: () => rPlanillaDia(dia) });
  const ir2 = (f) => { dia = f; mes = mesDe(f); history.replaceState(null, '', '#/metrados'); refrescar(); };
  box.querySelector('#d-fecha').onchange = (e) => e.target.value && ir2(e.target.value);
  box.querySelector('#d-ant').onclick = () => ir2(sumarDias(dia, -1));
  box.querySelector('#d-sig').onclick = () => ir2(sumarDias(dia, 1));
  box.querySelector('#d-add').onclick = async () => {
    const its = await elegirPartidas({ titulo: 'Partidas trabajadas el ' + fFecha(dia) });
    if (!its.length) return;
    if (!extraDia.has(dia)) extraDia.set(dia, new Set());
    its.forEach((it) => extraDia.get(dia).add(it));
    refrescar().then(() => { const i = document.querySelector('#t-dia tr[data-it] input[data-met]'); i?.focus(); i?.select(); });
  };
  box.querySelector('#d-calc')?.addEventListener('click', () => {
    o.calculadora = o.calculadora || { lineas: [], ajustes: {}, modo: o.parametros.redondeo || 'arriba', paso: 1, mo: false };
    o.calculadora.lineas = lineas.map((l) => ({ ...l }));
    o.calculadora.ajustes = {};
    cambio(); ir('calculadora');
  });
  box.querySelector('#d-sal')?.addEventListener('click', () => ir('almacen', { tab: 'generar', desde: dia, hasta: dia }));
  box.querySelector('#d-par')?.addEventListener('click', () => ir('equipos', { tab: 'generar', desde: dia, hasta: dia }));
  const tb = box.querySelector('#t-dia tbody');
  tb?.addEventListener('change', (e) => {
    const el = e.target.closest('[data-met]');
    if (!el) return;
    setMetrado(dia, el.closest('tr').dataset.it, num(el.value));
    refrescar();
  });
  tb?.addEventListener('click', async (e) => {
    const tr = e.target.closest('tr[data-it]');
    if (!tr) return;
    const it = tr.dataset.it;
    if (e.target.closest('[data-del]')) { extraDia.get(dia)?.delete(it); setMetrado(dia, it, 0); refrescar(); }
    if (e.target.closest('[data-det]')) detalleMetrado(it);
  });
}

/** Planilla de metrados con detalle por elemento: N° veces × N° elementos × largo × ancho × alto. */
function detalleMetrado(it) {
  const o = S.obra, p = S.motor.porItem.get(it);
  const k = dia + '|' + it;
  const filas = JSON.parse(JSON.stringify(o.detalles[k] || [{ elemento: '', veces: 1, elem: 1, largo: null, ancho: null, alto: null }]));
  const parcial = (f) => { const ds = [f.largo, f.ancho, f.alto].filter((x) => x !== null && x !== '' && x !== undefined); return r2((Number(f.veces) || 0) * (Number(f.elem) || 0) * ds.reduce((a, b) => a * Number(b), 1) * 1e4) / 1e4; };
  const pinta = (raiz) => {
    raiz.querySelector('#det-b').innerHTML = filas.map((f, i) => `<tr data-i="${i}"><td><input data-k="elemento" value="${esc(f.elemento || '')}" placeholder="Eje A-B, zapata Z-1…"></td>${['veces', 'elem', 'largo', 'ancho', 'alto'].map((c) => `<td><input class="n" data-k="${c}" value="${f[c] ?? ''}" style="width:72px"></td>`).join('')}<td class="n"><b>${cant(parcial(f), 4)}</b></td><td><button class="btn-icono peligro" data-x="${i}">${icono('borrar')}</button></td></tr>`).join('');
    raiz.querySelector('#det-t').textContent = cant(filas.reduce((s, f) => s + parcial(f), 0), 4);
  };
  ventana({
    titulo: `Planilla de metrados · ${it} · ${fFecha(dia)}`, ancho: '900px',
    cuerpo: `<p style="margin-top:0"><b>${esc(p.descripcion)}</b> (${esc(p.und)})</p><p class="tenue" style="font-size:12px">Parcial = N° veces × N° elementos × largo × ancho × alto (deje en blanco las medidas que no corresponden).</p>
      <div class="tabla-cont baja"><table class="t"><thead><tr><th>Elemento / descripción</th><th class="n">N° veces</th><th class="n">N° elem.</th><th class="n">Largo</th><th class="n">Ancho</th><th class="n">Alto</th><th class="n">Parcial</th><th></th></tr></thead><tbody id="det-b"></tbody>
      <tfoot><tr class="tot"><td colspan="6">METRADO DEL DÍA (${esc(p.und)})</td><td class="n" id="det-t"></td><td></td></tr></tfoot></table></div>
      <button class="btn btn-sm" id="det-add" style="margin-top:8px">${icono('mas')}Agregar fila</button>`,
    alAbrir: (raiz) => {
      pinta(raiz);
      raiz.querySelector('#det-add').onclick = () => { filas.push({ elemento: '', veces: 1, elem: 1, largo: null, ancho: null, alto: null }); pinta(raiz); };
      raiz.querySelector('#det-b').addEventListener('change', (e) => { const el = e.target.closest('[data-k]'); if (!el) return; const f = filas[Number(el.closest('tr').dataset.i)]; f[el.dataset.k] = el.dataset.k === 'elemento' ? el.value : el.value.trim() === '' ? null : num(el.value); pinta(raiz); });
      raiz.querySelector('#det-b').addEventListener('click', (e) => { const b = e.target.closest('[data-x]'); if (!b) return; filas.splice(Number(b.dataset.x), 1); pinta(raiz); });
    },
    botones: [{ txt: 'Cancelar' }, { txt: 'Quitar detalle', clase: 'btn-peligro', accion: () => { delete o.detalles[k]; cambio(); refrescar(); } }, { txt: 'Guardar metrado', clase: 'btn-primario', icono: 'guardar', accion: () => {
      const val = filas.filter((f) => parcial(f) !== 0);
      const tot = r2(val.reduce((s, f) => s + parcial(f), 0) * 1e4) / 1e4;
      if (val.length) o.detalles[k] = val; else delete o.detalles[k];
      setMetrado(dia, it, tot);
      aviso(`Metrado del día: ${cant(tot, 4)} ${p.und}`);
      refrescar();
    } }],
  });
}

// =================================================================== INSUMOS DEL MES
function vistaInsumos(box) {
  const o = S.obra, m = S.motor;
  const ini = mes + '-01', fin = finDeMes(mes);
  const lAnt = [], lMes = [];
  for (const p of m.partidasP) {
    const a = (Number(o.anterior_inicial?.[p.item]) || 0) + m.sumaDiaria(p.item, null, sumarDias(ini, -1));
    const x = m.sumaDiaria(p.item, ini, fin);
    if (a) lAnt.push({ item: p.item, metrado: a });
    if (x) lMes.push({ item: p.item, metrado: x });
  }
  const cAnt = new Map(m.consumo(lAnt, { incluirMO: false }).map((x) => [normal(x.insumo), x.cantidad]));
  const cMes = m.consumo(lMes, { incluirMO: false });
  const { ins } = m.disponibilidad();
  const ingresado = new Map();
  for (const mv of o.movimientos) {
    if (mv.tipo !== 'INGRESO' || mv.fecha > fin) continue;
    const a = m.itemAlmacen(mv.codigo);
    if (a?.insumo_acu) { const k = normal(a.insumo_acu); ingresado.set(k, (ingresado.get(k) || 0) + (Number(mv.cantidad) || 0) * (Number(a.factor) || 1)); }
  }
  for (const e of o.equipos) if (e.insumo_acu && (e.inicio || e.os_fecha) && (e.inicio || e.os_fecha) <= fin) { const k = normal(e.insumo_acu); ingresado.set(k, (ingresado.get(k) || 0) + (Number(e.cantidad) || 0)); }
  const filas = cMes.map((x) => {
    const k = normal(x.insumo);
    const ctrl = m.controlaInsumo({ tipo: x.tipo, descripcion: x.insumo });
    const d = ins.get(k);
    const acu = (cAnt.get(k) || 0) + x.cantidad;
    const ing = ingresado.get(k) || 0;
    let est = 'CONFORME', cls = 'badge-ok';
    if (!ctrl) { est = 'NO SE CONTROLA'; cls = ''; }
    else if (!d?.any) { est = 'SIN INGRESO NI O/S'; cls = 'badge-error'; }
    else if (ing - acu < -0.005) { est = 'INGRESO INSUFICIENTE'; cls = 'badge-alerta'; }
    else if (!d.conf) { est = 'POR CONFIRMAR'; cls = 'badge-alerta'; }
    return { ...x, acu, ing, saldo: ing - acu, primer: d?.any, est, cls, ctrl };
  });
  const nSin = filas.filter((x) => x.est === 'SIN INGRESO NI O/S').length;
  box.innerHTML = `
    <div class="filtros"><select id="f-mes">${mesesPosibles().map((x) => `<option value="${x}" ${x === mes ? 'selected' : ''}>${nombreMes(x)}</option>`).join('')}</select>
      <span class="tenue">Consumo según A.C.U. × metrados frente a lo ingresado al almacén (en unidades del A.C.U., con el factor de cada ítem) y a las O/S de equipos.</span></div>
    ${nSin ? `<div class="nota nota-error">${nSin} insumo(s) se consumen sin ingreso ni O/S: las partidas que los usan pueden quedar en NO VALORIZAR. Registre el ingreso en el almacén (y enlácelo al insumo del A.C.U.).</div>` : ''}
    <div class="tabla-cont"><table class="t"><thead><tr><th>Insumo</th><th>Und</th><th>Tipo</th><th class="n">Consumo del mes</th><th class="n">Consumo acumulado</th><th class="n">Ingresado</th><th class="n">Saldo</th><th>Primer ingreso</th><th>Estado</th></tr></thead><tbody>
    ${filas.map((x) => `<tr><td>${esc(x.insumo)}</td><td>${esc(x.und)}</td><td><span class="badge">${x.tipo}</span></td><td class="n">${cant(x.cantidad, 2)}</td><td class="n">${cant(x.acu, 2)}</td><td class="n">${x.ctrl ? cant(x.ing, 2) : ''}</td><td class="n ${x.ctrl && x.saldo < -0.005 ? 'neg' : ''}">${x.ctrl ? cant(x.saldo, 2) : ''}</td><td>${fFecha(x.primer)}</td><td><span class="badge ${x.cls}">${x.est}</span></td></tr>`).join('') || '<tr><td colspan="9" class="vacio">Sin consumo en el mes.</td></tr>'}
    </tbody></table></div>`;
  box.querySelector('#f-mes').onchange = (e) => { mes = e.target.value; refrescar(); };
}

async function importar() {
  const f = await elegirArchivo('.xlsx,.xlsm,.csv');
  if (!f) return;
  try {
    const hojas = await trabajando('Leyendo…', () => leerLibro(f));
    const sumar = await confirmar('Si ya existe metrado para una partida en esa fecha: <b>Aceptar</b> = sumar · <b>Cancelar</b> = reemplazar.', { si: 'Sumar' });
    const r = importarMetrados(S.obra, hojas[0].filas, { sumar });
    cambio({ guardarYa: true });
    aviso(`${r.n} metrado(s) importado(s).${r.errores.length ? ' ' + r.errores.length + ' fila(s) con error: ' + r.errores.slice(0, 3).join('; ') : ''}`, r.errores.length ? 'alerta' : 'ok', 9000);
    refrescar();
  } catch (e) { aviso(e.message, 'error', 8000); }
}
