// Valorización mensual: periodos, 4.4 por partida, 4.3 resumen, 4.5 comparado, curva S, cierre del periodo.
import { S, cambio } from '../store.js';
import { esc, fmt, pct, cant, num, normal, fFecha, nombreMes, mesDe, r2 } from '../util.js';
import { icono, formulario, confirmar, aviso, ventana, elegirArchivo } from '../ui.js';
import { periodoSiguiente, periodoInicial } from '../calc.js';
import { rValorizacion, rResumen, rComparado, rCurvaS } from '../reportes.js';
import { cab, btnExport, enlazarExport, opcionesPeriodo, sinPresupuesto } from './comun.js';
import { refrescar } from '../app.js';

let k = null, tab = 'v44', q = '', soloAv = true, sub = '';
let grafico = null;
export function destruir() { grafico?.destroy(); grafico = null; }

export function render(cont) {
  const o = S.obra, m = S.motor;
  if (!m.partidasP.length) { cont.innerHTML = cab('Valorización mensual', '') + sinPresupuesto(); return; }
  if (!o.periodos.length) {
    cont.innerHTML = cab('Valorización mensual', '') + `<div class="panel"><div class="vacio"><h3>Sin periodos de valorización</h3><p>Cree la valorización N° 01 (desde la fecha de inicio hasta fin de ese mes).</p><button class="btn btn-primario" id="b-p1">${icono('mas')}Crear valorización N° 01</button></div></div>`;
    cont.querySelector('#b-p1').onclick = () => { const p = periodoInicial(o); if (!p) { aviso('Complete la fecha de inicio en Datos de la obra.', 'alerta'); return; } o.periodos.push(p); cambio(); refrescar(); };
    return;
  }
  const pers = m.periodos();
  if (k === null || k >= pers.length) { const a = pers.findIndex((p) => p.estado !== 'CERRADO'); k = a >= 0 ? a : pers.length - 1; }
  const per = pers[k];
  const v = m.valorizacion(k);
  const c = m.comparado(k);
  const abierto = per.estado !== 'CERRADO';
  const ultimo = k === pers.length - 1;
  const nMan = Object.keys(per.manual || {}).length;
  cont.innerHTML = `
    ${cab('Valorización mensual', `Valorización N° <b>${String(per.numero).padStart(2, '0')}</b> · del ${fFecha(per.del)} al ${fFecha(per.al)} · ${nombreMes(mesDe(per.al))} · ${abierto ? '<span class="badge badge-ok">ABIERTA</span>' : '<span class="badge">CERRADA</span>'}`,
      `<select id="sel-per">${opcionesPeriodo(k)}</select>
       ${abierto ? `<button class="btn" id="b-edit">${icono('editar')}Periodo</button><button class="btn btn-oro" id="b-cerrar">${icono('candado')}Cerrar valorización</button>` : `${ultimo ? `<button class="btn" id="b-reabrir">Reabrir</button>` : ''}`}
       ${ultimo && !abierto ? `<button class="btn btn-primario" id="b-nuevo">${icono('mas')}Siguiente valorización</button>` : ''}`)}
    <div class="tarjetas">
      <div class="tarjeta"><div class="et">C.D. valorizado del mes</div><div class="val">S/ ${fmt(v.cd.act)}</div><div class="sub">${pct(v.cd.pctAct)} del C.D. · ${v.nPartidasMes} partidas</div></div>
      <div class="tarjeta"><div class="et">Costos indirectos del mes</div><div class="val">S/ ${fmt(v.totalIndirecto.act)}</div><div class="sub">${v.ind.length} conceptos</div></div>
      <div class="tarjeta"><div class="et">Total del mes</div><div class="val">S/ ${fmt(v.total.act)}</div><div class="sub">Acumulado S/ ${fmt(v.total.acu)}</div></div>
      <div class="tarjeta"><div class="et">Avance físico acumulado</div><div class="val">${pct(v.cd.pctAcu)}</div><div class="barra-prog"><div style="width:${Math.min(100, v.cd.pctAcu * 100)}%"></div></div></div>
      <div class="tarjeta"><div class="et">Programado del mes</div><div class="val">S/ ${fmt(c.programado)}</div><div class="sub">Cumplimiento ${c.cumplimiento !== null ? pct(c.cumplimiento) : '—'}</div></div>
    </div>
    ${nMan ? `<div class="nota nota-alerta">${nMan} partida(s) tienen el metrado del mes digitado a mano (celdas amarillas) en lugar de la suma de los metrados diarios.</div>` : ''}
    <div class="pestanas">${[['v44', '4.4 Valorización por partida'], ['v43', '4.3 Resumen'], ['v45', '4.5 Programado vs ejecutado'], ['curva', 'Curva S'], ['ant', 'Acumulado anterior al sistema']].map(([x, t]) => `<button data-tab="${x}" class="${tab === x ? 'activo' : ''}">${t}</button>`).join('')}</div>
    <div id="cuerpo"></div>`;
  cont.querySelector('#sel-per').onchange = (e) => { k = Number(e.target.value); refrescar(); };
  cont.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { tab = b.dataset.tab; refrescar(); }));
  cont.querySelector('#b-edit')?.addEventListener('click', () => editarPeriodo(per));
  cont.querySelector('#b-cerrar')?.addEventListener('click', () => cerrar(per));
  cont.querySelector('#b-reabrir')?.addEventListener('click', async () => {
    const sig = pers[k + 1];
    if (sig) { aviso('Solo se puede reabrir la última valorización.', 'alerta'); return; }
    if (!(await confirmar('¿Reabrir esta valorización? El metrado del mes vuelve a calcularse con los metrados diarios (se pierden los valores congelados al cierre, salvo los digitados a mano).'))) return;
    per.estado = 'ABIERTO'; per.cerrado = null; delete per.fecha_cierre;
    cambio(); refrescar();
  });
  cont.querySelector('#b-nuevo')?.addEventListener('click', () => { o.periodos.push(periodoSiguiente(o, per)); cambio(); k = k + 1; aviso('Nueva valorización creada.'); refrescar(); });
  const box = cont.querySelector('#cuerpo');
  ({ v44, v43, v45, curva, ant })[tab](box, v, c, per, abierto);
}

function v44(box, v, c, per, abierto) {
  const m = S.motor;
  const nq = normal(q);
  let filas = v.filas;
  if (sub) filas = filas.filter((f) => m.raiz(f.item) === sub);
  if (soloAv || nq) {
    const keep = new Set();
    for (const f of filas) if (f.tipo === 'P' && (!soloAv || f.act.m || f.ant.m || per.manual?.[f.item] !== undefined) && (!nq || normal(f.item + ' ' + f.descripcion).includes(nq))) { keep.add(f.item); let pa = m.padre.get(f.item); while (pa) { keep.add(pa); pa = m.padre.get(pa); } }
    filas = filas.filter((f) => keep.has(f.item));
  }
  const td = (x, d = 2) => `<td class="n">${fmt(x, d, true)}</td>`;
  box.innerHTML = `
    <div class="filtros"><input type="search" class="buscar" id="q" placeholder="Buscar partida…" value="${esc(q)}">
      <select id="f-sub"><option value="">Todos los subpresupuestos</option>${m.raices.map((it) => `<option value="${esc(it)}" ${sub === it ? 'selected' : ''}>${esc(it)} ${esc(m.porItem.get(it).descripcion.slice(0, 40))}</option>`).join('')}</select>
      <label class="chk"><input type="checkbox" id="f-av" ${soloAv ? 'checked' : ''}> Solo partidas con avance</label><span class="tenue" style="font-size:12px">${filas.length} filas</span><span style="flex:1"></span>${btnExport('v44', '4.4')}</div>
    <div class="tabla-cont"><table class="t" id="t-val"><thead>
      <tr><th rowspan="2">Ítem</th><th rowspan="2">Descripción</th><th rowspan="2">Und</th><th colspan="3" class="c">Presupuesto</th><th colspan="3" class="c">Acumulado anterior</th><th colspan="3" class="c">Del mes (actual)</th><th colspan="3" class="c">Acumulado actual</th><th colspan="2" class="c">Saldo</th></tr>
      <tr><th class="n">Metrado</th><th class="n">P.U.</th><th class="n">Parcial</th><th class="n">Metrado</th><th class="n">S/</th><th class="n">%</th><th class="n">Metrado</th><th class="n">S/</th><th class="n">%</th><th class="n">Metrado</th><th class="n">S/</th><th class="n">%</th><th class="n">Metrado</th><th class="n">S/</th></tr></thead><tbody>
      ${filas.map((f) => f.tipo === 'T'
        ? `<tr class="tit ${f.nivel === 1 ? 'tit-1' : ''}"><td class="item">${esc(f.item)}</td><td class="desc"><span class="sangria-${Math.min(f.nivel, 7)}">${esc(f.descripcion)}</span></td><td></td><td></td><td></td>${td(f.parcial)}<td></td>${td(f.ant.s)}<td class="n">${pct(f.parcial ? f.ant.s / f.parcial : 0)}</td><td></td>${td(f.act.s)}<td class="n">${pct(f.pctAct)}</td><td></td>${td(f.acu.s)}<td class="n">${pct(f.pctAcu)}</td><td></td>${td(f.sal.s)}</tr>`
        : `<tr data-it="${esc(f.item)}" class="${f.excede ? 'e-met' : ''}"><td class="item">${esc(f.item)}</td><td class="desc"><span class="sangria-${Math.min(f.nivel, 7)}">${esc(f.descripcion)}</span>${f.excede ? ' <span class="badge badge-error">supera E.T.</span>' : ''}</td><td>${esc(f.und)}</td>${td(f.metrado)}${td(f.pu)}${td(f.parcial)}${td(f.ant.m)}${td(f.ant.s)}<td class="n">${pct(f.parcial ? f.ant.s / f.parcial : 0)}</td>
          <td>${abierto ? `<input class="n ${f.fuente === 'manual' ? 'editado' : ''}" data-act value="${f.act.m ? cant(f.act.m, 4) : ''}" style="width:84px" title="${f.fuente === 'manual' ? 'Digitado a mano (Σ diario: ' + cant(S.motor.sumaDiaria(f.item, per.del, per.al), 4) + '). Borre para volver a la suma diaria.' : 'Suma de los metrados diarios del periodo'}">` : `<span class="n" style="display:block;text-align:right">${fmt(f.act.m, 2, true)}</span>`}</td>${td(f.act.s)}<td class="n">${pct(f.pctAct)}</td>${td(f.acu.m)}${td(f.acu.s)}<td class="n">${pct(f.pctAcu)}</td>${td(f.sal.m)}${td(f.sal.s)}</tr>`).join('')}
    </tbody><tfoot><tr class="tot"><td></td><td>COSTO DIRECTO</td><td></td><td></td><td></td>${td(v.cd.presupuesto)}<td></td>${td(v.cd.ant)}<td class="n">${pct(v.cd.presupuesto ? v.cd.ant / v.cd.presupuesto : 0)}</td><td></td>${td(v.cd.act)}<td class="n">${pct(v.cd.pctAct)}</td><td></td>${td(v.cd.acu)}<td class="n">${pct(v.cd.pctAcu)}</td><td></td>${td(v.cd.sal)}</tr></tfoot></table></div>
    ${abierto ? '<p class="tenue" style="font-size:12px">El metrado del mes es la suma de los metrados diarios del periodo. Puede digitarlo directamente (por ejemplo, si no lleva el control diario); para volver a la suma diaria, borre el valor.</p>' : ''}`;
  enlazarExport(box, { v44: () => rValorizacion(v.k, { soloConAvance: false }) });
  let t;
  box.querySelector('#q').oninput = (e) => { clearTimeout(t); t = setTimeout(() => { q = e.target.value; refrescar().then(() => { const el = document.querySelector('#q'); el?.focus(); el?.setSelectionRange(el.value.length, el.value.length); }); }, 300); };
  box.querySelector('#f-av').onchange = (e) => { soloAv = e.target.checked; refrescar(); };
  box.querySelector('#f-sub').onchange = (e) => { sub = e.target.value; refrescar(); };
  box.querySelector('#t-val tbody').addEventListener('change', (e) => {
    const el = e.target.closest('[data-act]');
    if (!el) return;
    const it = el.closest('tr').dataset.it;
    per.manual = per.manual || {};
    if (el.value.trim() === '') delete per.manual[it];
    else per.manual[it] = num(el.value);
    cambio(); refrescar();
  });
}

function v43(box, v) {
  const fila = (cpt, x, cls = '', item = '') => `<tr class="${cls}"><td class="item">${esc(item)}</td><td>${esc(cpt)}</td><td class="n">${fmt(x.presupuesto)}</td><td class="n">${fmt(x.ant)}</td><td class="n">${fmt(x.act)}</td><td class="n">${pct(x.pctAct)}</td><td class="n">${fmt(x.acu)}</td><td class="n">${pct(x.pctAcu)}</td><td class="n">${fmt(x.sal)}</td></tr>`;
  const f10 = { ET: 'Metrados del expediente técnico', MAYOR_METRADO: 'Mayores metrados', PARTIDA_NUEVA: 'Partidas nuevas', DEDUCTIVO: 'Deductivos' };
  box.innerHTML = `<div class="filtros"><span style="flex:1"></span>${btnExport('v43', '4.3')}</div>
    <div class="panel"><div class="tabla-cont libre" style="border:0"><table class="t"><thead><tr><th>Ítem</th><th>Concepto</th><th class="n">Presupuesto S/</th><th class="n">Anterior S/</th><th class="n">Del mes S/</th><th class="n">% mes</th><th class="n">Acumulado S/</th><th class="n">% acum.</th><th class="n">Saldo S/</th></tr></thead><tbody>
      ${v.subp.map((s) => fila(s.descripcion, s, '', s.item)).join('')}
      ${fila('COSTO DIRECTO', v.cd, 'tit tit-1')}
      ${v.ind.map((x) => fila(`${x.concepto} (${fmt(x.ratio * 100)} % C.D.)`, x)).join('')}
      ${fila('TOTAL COSTO INDIRECTO', v.totalIndirecto, 'tit')}
      ${fila('PRESUPUESTO TOTAL', v.total, 'tot')}
    </tbody></table></div></div>
    <div class="panel"><div class="panel-cab"><h2>Formato N° 10 (4.3.1): E.T. + mayores metrados + partidas nuevas − deductivos</h2></div><div class="tabla-cont libre" style="border:0"><table class="t"><thead><tr><th>Concepto</th><th class="n">Presupuesto S/</th><th class="n">Anterior S/</th><th class="n">Del mes S/</th><th class="n">Acumulado S/</th></tr></thead><tbody>
      ${Object.entries(v.f10).map(([o, x]) => `<tr><td>${f10[o] || o}</td><td class="n">${fmt(x.presupuesto)}</td><td class="n">${fmt(x.ant)}</td><td class="n">${fmt(x.act)}</td><td class="n">${fmt(x.ant + x.act)}</td></tr>`).join('')}
    </tbody></table></div><div class="panel-cuerpo"><p class="tenue" style="font-size:12px;margin:0">El origen de cada partida (E.T., mayor metrado, partida nueva o deductivo) se define al editarla en el presupuesto.</p></div></div>`;
  enlazarExport(box, { v43: () => rResumen(v.k) });
}

function v45(box, v, c) {
  const filas = c.filas.filter((f) => f.tipo === 'P' && (f.programado || f.ejecutado));
  box.innerHTML = `<div class="filtros"><span class="tenue">Programado de ${nombreMes(c.ym)} según el cronograma valorizado vigente.</span><span style="flex:1"></span>${btnExport('v45', '4.5')}</div>
    ${!S.motor.cronoMeses().length ? '<div class="nota nota-alerta">No hay cronograma valorizado. <a href="#/cronograma">Cárguelo o genérelo</a>.</div>' : ''}
    <div class="tarjetas"><div class="tarjeta"><div class="et">Programado</div><div class="val">S/ ${fmt(c.programado)}</div></div><div class="tarjeta"><div class="et">Ejecutado</div><div class="val">S/ ${fmt(c.ejecutado)}</div></div><div class="tarjeta"><div class="et">Diferencia</div><div class="val ${c.diferencia < 0 ? 'neg' : 'pos'}">S/ ${fmt(c.diferencia)}</div></div><div class="tarjeta"><div class="et">% de cumplimiento</div><div class="val">${c.cumplimiento !== null ? pct(c.cumplimiento) : '—'}</div></div></div>
    <div class="tabla-cont"><table class="t"><thead><tr><th>Ítem</th><th>Partida</th><th class="n">Programado S/</th><th class="n">Ejecutado S/</th><th class="n">Diferencia S/</th><th class="n">% cumplimiento</th></tr></thead><tbody>
      ${filas.map((f) => `<tr><td class="item">${esc(f.item)}</td><td class="desc">${esc(f.descripcion)}</td><td class="n">${fmt(f.programado)}</td><td class="n">${fmt(f.ejecutado)}</td><td class="n ${f.diferencia < 0 ? 'neg' : 'pos'}">${fmt(f.diferencia)}</td><td class="n">${f.cumplimiento !== null ? pct(f.cumplimiento) : '—'}</td></tr>`).join('')}
    </tbody><tfoot><tr class="tot"><td></td><td>COSTO DIRECTO</td><td class="n">${fmt(c.programado)}</td><td class="n">${fmt(c.ejecutado)}</td><td class="n">${fmt(c.diferencia)}</td><td class="n">${c.cumplimiento !== null ? pct(c.cumplimiento) : ''}</td></tr></tfoot></table></div>`;
  enlazarExport(box, { v45: () => rComparado(v.k) });
}

function curva(box) {
  const o = S.obra, m = S.motor;
  const cs = m.curvaS();
  box.innerHTML = `<div class="filtros"><span class="tenue">El devengado (avance financiero) se digita por mes como monto acumulado (del SIAF).</span><span style="flex:1"></span>${btnExport('cs', 'Curva S')}</div>
    <div class="panel"><div class="panel-cuerpo"><div class="graf" style="height:340px"><canvas id="g-cs"></canvas></div></div></div>
    <div class="tabla-cont libre"><table class="t" id="t-cs"><thead><tr><th>N°</th><th>Mes</th><th class="n">Programado S/</th><th class="n">Prog. acumulado</th><th class="n">% prog.</th><th class="n">Ejecutado S/</th><th class="n">Ejec. acumulado</th><th class="n">% físico</th><th class="n">Devengado acumulado S/</th><th class="n">% financiero</th></tr></thead><tbody>
    ${cs.map((x) => `<tr><td>${x.n}</td><td>${nombreMes(x.ym)}</td><td class="n">${fmt(x.programado)}</td><td class="n">${fmt(x.progAcum)}</td><td class="n">${pct(x.progPct)}</td><td class="n">${x.ejecutado !== null ? fmt(x.ejecutado) : ''}</td><td class="n">${x.ejecAcum !== null ? fmt(x.ejecAcum) : ''}</td><td class="n">${x.ejecPct !== null ? pct(x.ejecPct) : ''}</td><td><input class="n" data-dev="${x.ym}" value="${x.devengado !== null ? fmt(x.devengado) : ''}" style="width:120px"></td><td class="n">${x.devPct !== null ? pct(x.devPct) : ''}</td></tr>`).join('')}
    </tbody></table></div>`;
  enlazarExport(box, { cs: () => rCurvaS() });
  box.querySelector('#t-cs').addEventListener('change', (e) => {
    const el = e.target.closest('[data-dev]');
    if (!el) return;
    if (el.value.trim() === '') delete o.financiero.devengado_acumulado[el.dataset.dev];
    else o.financiero.devengado_acumulado[el.dataset.dev] = num(el.value);
    cambio(); refrescar();
  });
  const c = document.getElementById('g-cs');
  if (!c || !window.Chart) return;
  destruir();
  grafico = new window.Chart(c, {
    data: {
      labels: cs.map((x) => nombreMes(x.ym, true)),
      datasets: [
        { type: 'bar', label: 'Programado del mes (S/)', data: cs.map((x) => x.programado), backgroundColor: 'rgba(44,82,130,.18)', yAxisID: 'y2' },
        { type: 'bar', label: 'Ejecutado del mes (S/)', data: cs.map((x) => x.ejecutado), backgroundColor: 'rgba(224,165,47,.55)', yAxisID: 'y2' },
        { type: 'line', label: '% programado acum.', data: cs.map((x) => +(x.progPct * 100).toFixed(2)), borderColor: '#2c5282', tension: .3, yAxisID: 'y' },
        { type: 'line', label: '% físico acum.', data: cs.map((x) => (x.ejecPct === null ? null : +(x.ejecPct * 100).toFixed(2))), borderColor: '#e0a52f', tension: .3, yAxisID: 'y' },
        { type: 'line', label: '% financiero acum.', data: cs.map((x) => (x.devPct === null ? null : +(x.devPct * 100).toFixed(2))), borderColor: '#1e8e5a', borderDash: [5, 4], tension: .3, yAxisID: 'y' },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } } }, scales: { y: { min: 0, max: 100, position: 'left', ticks: { callback: (x) => x + ' %' } }, y2: { position: 'right', grid: { display: false }, ticks: { callback: (x) => 'S/ ' + Number(x).toLocaleString('en-US') } } } },
  });
}

function ant(box) {
  const o = S.obra, m = S.motor;
  const conAnt = m.partidasP.filter((p) => Number(o.anterior_inicial[p.item]));
  box.innerHTML = `<div class="nota">Si la obra ya tenía valorizaciones antes de usar este programa, registre aquí el <b>metrado acumulado anterior</b> de cada partida (se suma como «anterior» a la primera valorización). Puede importarlo desde el JSON de una valorización o desde Excel (columnas ITEM y METRADO).</div>
    <div class="filtros"><button class="btn btn-sm" id="b-xl">${icono('importar')}Importar desde Excel</button><button class="btn btn-sm" id="b-add">${icono('mas')}Agregar partida</button><span style="flex:1"></span><span class="badge badge-azul">S/ ${fmt(r2(conAnt.reduce((s, p) => s + r2(o.anterior_inicial[p.item] * p.pu), 0)))}</span></div>
    <div class="tabla-cont"><table class="t" id="t-ant"><thead><tr><th>Ítem</th><th>Partida</th><th>Und</th><th class="n">Metrado E.T.</th><th class="n">Acumulado anterior</th><th class="n">S/</th><th></th></tr></thead><tbody>
    ${conAnt.map((p) => `<tr data-it="${esc(p.item)}"><td class="item">${esc(p.item)}</td><td class="desc">${esc(p.descripcion)}</td><td>${esc(p.und)}</td><td class="n">${fmt(p.metrado)}</td><td><input class="n" data-ant value="${cant(o.anterior_inicial[p.item], 4)}" style="width:100px"></td><td class="n">${fmt(r2(o.anterior_inicial[p.item] * p.pu))}</td><td><button class="btn-icono peligro" data-del>${icono('borrar')}</button></td></tr>`).join('') || '<tr><td colspan="7" class="vacio">Sin acumulado anterior (la obra empieza con este programa).</td></tr>'}
    </tbody></table></div>`;
  const tb = box.querySelector('#t-ant tbody');
  tb.addEventListener('change', (e) => { const el = e.target.closest('[data-ant]'); if (!el) return; const it = el.closest('tr').dataset.it; const v = num(el.value); if (v) o.anterior_inicial[it] = v; else delete o.anterior_inicial[it]; cambio(); refrescar(); });
  tb.addEventListener('click', (e) => { if (!e.target.closest('[data-del]')) return; delete o.anterior_inicial[e.target.closest('tr').dataset.it]; cambio(); refrescar(); });
  box.querySelector('#b-add').onclick = async () => {
    const { elegirPartidas } = await import('./comun.js');
    const its = await elegirPartidas({ titulo: 'Partidas con acumulado anterior' });
    its.forEach((it) => { if (!o.anterior_inicial[it]) o.anterior_inicial[it] = Number(m.porItem.get(it).metrado) || 1; });
    if (its.length) { cambio(); refrescar(); }
  };
  box.querySelector('#b-xl').onclick = async () => {
    const f = await elegirArchivo('.xlsx,.xlsm,.csv');
    if (!f) return;
    const { leerLibro, leerTabla, claveItem } = await import('../importar.js');
    try {
      const h = await leerLibro(f);
      const t = leerTabla(h[0].filas, { item: (x) => x === 'ITEM' || x === 'PARTIDA', metrado: (x) => x.startsWith('METRADO') || x.startsWith('ACUM') || x.startsWith('CANT') }, 2);
      if (!t) throw new Error('Se esperan columnas ITEM y METRADO.');
      const porClave = new Map(m.partidasP.map((p) => [claveItem(p.item), p.item]));
      let n = 0;
      t.filas.forEach((r) => { const it = porClave.get(claveItem(String(r.item ?? '').trim())); const v = num(r.metrado); if (it && v) { o.anterior_inicial[it] = v; n++; } });
      cambio(); aviso(`${n} acumulado(s) anterior(es) importado(s).`); refrescar();
    } catch (e) { aviso(e.message, 'error', 8000); }
  };
}

function editarPeriodo(per) {
  formulario({
    titulo: 'Periodo de la valorización', ancho: '560px',
    campos: [{ id: 'numero', et: 'N° de valorización', tipo: 'number', req: true }, { id: 'del', et: 'Del', tipo: 'date', req: true }, { id: 'al', et: 'Al', tipo: 'date', req: true }],
    val: per,
    alGuardar: (v) => {
      if (v.al < v.del) throw new Error('La fecha final es anterior a la inicial.');
      const otros = S.obra.periodos.filter((p) => p !== per);
      if (otros.some((p) => !(v.al < p.del || v.del > p.al))) throw new Error('El periodo se cruza con otra valorización.');
      Object.assign(per, v); cambio(); refrescar();
    },
  });
}
async function cerrar(per) {
  const m = S.motor;
  const k2 = m.periodos().indexOf(per);
  const v = m.valorizacion(k2);
  const exc = v.filas.filter((f) => f.excede).length;
  const est = m.estadosMes(mesDe(per.al));
  const nNo = [...est.values()].filter((e) => e.estado.cod === 'NO_VALORIZAR').length;
  ventana({
    titulo: `Cerrar la valorización N° ${String(per.numero).padStart(2, '0')}`, ancho: '600px',
    cuerpo: `<p>Al cerrar, el metrado del mes de cada partida queda <b>congelado</b> y pasa como «acumulado anterior» a la siguiente valorización (como CERRAR MES de los programas en Excel).</p>
      <div class="kv"><div>C.D. del mes</div><div><b>S/ ${fmt(v.cd.act)}</b> (${pct(v.cd.pctAct)})</div><div>Total con indirectos</div><div><b>S/ ${fmt(v.total.act)}</b></div><div>Acumulado</div><div>S/ ${fmt(v.total.acu)} · ${pct(v.cd.pctAcu)} físico</div></div>
      ${exc ? `<div class="nota nota-alerta">${exc} partida(s) superan el metrado del expediente.</div>` : ''}${nNo ? `<div class="nota nota-error">${nNo} partida(s) tienen días marcados «NO VALORIZAR» en el control diario. Revíselas antes de cerrar.</div>` : ''}
      <p class="tenue" style="font-size:12px">Antes de cerrar, exporte los reportes del mes y descargue un respaldo.</p>`,
    botones: [{ txt: 'Cancelar' }, { txt: 'Cerrar y crear la siguiente', clase: 'btn-oro', icono: 'candado', accion: () => {
      const congelado = {};
      for (const p of m.partidasP) { const x = m.metradoPeriodo(per, p.item).m; if (x) congelado[p.item] = x; }
      per.cerrado = congelado; per.estado = 'CERRADO'; per.fecha_cierre = new Date().toISOString();
      const o = S.obra;
      const ultimo = m.periodos()[m.periodos().length - 1] === per;
      if (ultimo && (!m.termino || per.al < m.termino)) o.periodos.push(periodoSiguiente(o, per));
      cambio({ guardarYa: true });
      k = null;
      aviso('Valorización cerrada.');
      refrescar();
    } }],
  });
}
