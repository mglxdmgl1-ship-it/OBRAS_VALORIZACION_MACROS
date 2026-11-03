// Panel de la obra: indicadores, curva S, resumen por subpresupuesto y alertas.
import { S } from '../store.js';
import { esc, fmt, pct, fFecha, hoy, nombreMes, mesDe } from '../util.js';
import { icono } from '../ui.js';
import { cab, sinPresupuesto } from './comun.js';

let grafico = null;
export function destruir() { grafico?.destroy(); grafico = null; }

export function render(cont) {
  const o = S.obra, m = S.motor, d = o.datos;
  if (!m.partidasP.length) { cont.innerHTML = cab(d.nombre_corto || 'Panel', esc(d.nombre)) + sinPresupuesto(); return; }
  const pers = m.periodos();
  const k = pers.length ? pers.findIndex((p) => p.estado !== 'CERRADO') : -1;
  const kk = k >= 0 ? k : pers.length - 1;
  const v = kk >= 0 ? m.valorizacion(kk) : null;
  const cmp = kk >= 0 ? m.comparado(kk) : null;
  const cs = m.curvaS();
  const ymV = v ? mesDe(v.periodo.al) : mesDe(hoy());
  const prog = cs.filter((x) => x.ym <= ymV).pop();
  const corte = v ? v.periodo.al : hoy();
  const plazo = m.avancePlazo(corte);
  const est = m.estadosMes(ymV);
  const nNo = [...est.values()].filter((e) => e.estado.cod === 'NO_VALORIZAR').length;
  const nRev = [...est.values()].filter((e) => e.estado.orden > 0 && e.estado.cod !== 'NO_VALORIZAR').length;
  const neg = [...m.stock().entries()].filter(([, x]) => x.saldo < -0.0001);
  const exced = v ? v.filas.filter((f) => f.excede) : [];
  const sinACU = m.partidasP.filter((p) => !m.acu(p.item).length).length;

  cont.innerHTML = `
    ${cab(d.nombre_corto || 'Panel de la obra', `${esc(d.nombre)}<br><span class="mono">CUI ${esc(d.cui || '—')} · Meta ${esc(d.meta || '—')}</span> · ${esc([d.distrito, d.provincia, d.region].filter(Boolean).join(' – '))} · ${esc(d.modalidad || '')}`,
      `<a class="btn" href="#/metrados">${icono('metrados')}Metrados del día</a><a class="btn btn-primario" href="#/valorizacion">${icono('valorizacion')}Valorización</a>`)}
    <div class="tarjetas">
      ${tarjeta('Presupuesto total', 'S/ ' + fmt(m.presupuestoTotal()), `C.D. S/ ${fmt(m.costoDirecto())} · C.I. S/ ${fmt(m.totalIndirecto())}`)}
      ${tarjeta('Valorizado acumulado (C.D.)', 'S/ ' + fmt(v?.cd.acu || 0), v ? `Al ${fFecha(v.periodo.al)} · con indirectos S/ ${fmt(v.total.acu)}` : 'Sin periodos', v?.cd.pctAcu)}
      ${tarjeta('Avance físico', pct(v?.cd.pctAcu || 0), `Programado al mes: ${prog ? pct(prog.progPct) : '—'}`, v?.cd.pctAcu)}
      ${tarjeta('Valorizado del mes', 'S/ ' + fmt(v?.cd.act || 0), v ? `Val. N° ${String(v.periodo.numero).padStart(2, '0')} · ${pct(v.cd.pctAct)} del C.D.` : '')}
      ${tarjeta('Cumplimiento del mes', cmp?.cumplimiento !== null && cmp ? pct(cmp.cumplimiento) : '—', cmp ? `Programado S/ ${fmt(cmp.programado)}` : 'Sin cronograma', cmp?.cumplimiento)}
      ${tarjeta('Plazo', plazo ? `${plazo.transcurridos} / ${plazo.total} d` : '—', plazo ? `${fFecha(m.inicio)} → ${fFecha(m.termino)}` : 'Complete fecha de inicio y plazo', plazo?.pct)}
    </div>
    <div class="rejilla-3">
      <div class="panel"><div class="panel-cab"><h2>Curva S</h2><span class="tenue" style="font-size:12px">% acumulado del C.D. (programado vs físico) y financiero</span></div><div class="panel-cuerpo"><div class="graf"><canvas id="g-curva"></canvas></div></div></div>
      <div class="panel"><div class="panel-cab"><h2>Alertas</h2></div><div class="panel-cuerpo">
        <ul class="lista-simple">
          ${alerta(nNo, `partida(s) con días para <b>NO VALORIZAR</b> en ${nombreMes(ymV)}`, '#/metrados', 'error')}
          ${alerta(nRev, 'partida(s) para revisar (insumo por confirmar, rendimiento o metrado)', '#/metrados', 'alerta')}
          ${alerta(exced.length, 'partida(s) cuyo acumulado supera el metrado del expediente', '#/valorizacion', 'alerta')}
          ${alerta(neg.length, 'ítem(s) del almacén con saldo negativo', '#/almacen', 'error')}
          ${alerta(sinACU, 'partida(s) sin A.C.U. (no calculan insumos)', '#/acu', 'alerta')}
          ${alerta(m.cronoMeses().length ? 0 : 1, 'Falta el cronograma valorizado (para el 4.5 y la curva S)', '#/cronograma', 'alerta')}
        </ul>
        ${!nNo && !nRev && !exced.length && !neg.length && !sinACU && m.cronoMeses().length ? '<p class="nota nota-ok">Sin alertas.</p>' : ''}
      </div></div>
    </div>
    <div class="panel"><div class="panel-cab"><h2>Resumen por subpresupuesto ${v ? `· Valorización N° ${String(v.periodo.numero).padStart(2, '0')}` : ''}</h2><a class="btn btn-sm" href="#/reportes">${icono('reportes')}Reportes</a></div>
      <div class="tabla-cont libre"><table class="t"><thead><tr><th>Ítem</th><th>Subpresupuesto</th><th class="n">Presupuesto S/</th><th class="n">Anterior S/</th><th class="n">Del mes S/</th><th class="n">Acumulado S/</th><th class="n">% acum.</th><th style="width:140px"></th><th class="n">Saldo S/</th></tr></thead><tbody>
      ${(v?.subp || m.subpresupuestos().map((s) => ({ ...s, ant: 0, act: 0, acu: 0, pctAcu: 0, sal: s.presupuesto }))).map((s) => `<tr><td class="item">${esc(s.item)}</td><td>${esc(s.descripcion)}</td><td class="n">${fmt(s.presupuesto)}</td><td class="n">${fmt(s.ant)}</td><td class="n">${fmt(s.act)}</td><td class="n">${fmt(s.acu)}</td><td class="n">${pct(s.pctAcu)}</td><td><div class="barra-prog"><div style="width:${Math.min(100, (s.pctAcu || 0) * 100)}%"></div></div></td><td class="n">${fmt(s.sal)}</td></tr>`).join('')}
      ${v ? `<tr class="tot"><td></td><td>COSTO DIRECTO</td><td class="n">${fmt(v.cd.presupuesto)}</td><td class="n">${fmt(v.cd.ant)}</td><td class="n">${fmt(v.cd.act)}</td><td class="n">${fmt(v.cd.acu)}</td><td class="n">${pct(v.cd.pctAcu)}</td><td></td><td class="n">${fmt(v.cd.sal)}</td></tr>` : ''}
      </tbody></table></div></div>`;
  dibujarCurva(cs);
}
function tarjeta(et, val, sub, p) {
  return `<div class="tarjeta"><div class="et">${et}</div><div class="val">${val}</div><div class="sub">${sub || ''}</div>${p !== undefined && p !== null ? `<div class="barra-prog"><div style="width:${Math.max(0, Math.min(100, p * 100))}%"></div></div>` : ''}</div>`;
}
function alerta(n, txt, href, tipo) {
  if (!n) return '';
  return `<li><a href="${href}"><span class="badge badge-${tipo}">${n}</span></a> ${txt}</li>`;
}
function dibujarCurva(cs) {
  const c = document.getElementById('g-curva');
  if (!c || !window.Chart || !cs.length) return;
  destruir();
  grafico = new window.Chart(c, {
    type: 'line',
    data: {
      labels: cs.map((x) => nombreMes(x.ym, true)),
      datasets: [
        { label: 'Programado acumulado', data: cs.map((x) => +(x.progPct * 100).toFixed(2)), borderColor: '#2c5282', backgroundColor: 'rgba(44,82,130,.08)', fill: true, tension: .3, borderWidth: 2.5, pointRadius: 3 },
        { label: 'Físico ejecutado acumulado', data: cs.map((x) => (x.ejecPct === null ? null : +(x.ejecPct * 100).toFixed(2))), borderColor: '#e0a52f', backgroundColor: '#e0a52f', tension: .3, borderWidth: 2.5, pointRadius: 4, spanGaps: false },
        { label: 'Financiero (devengado)', data: cs.map((x) => (x.devPct === null ? null : +(x.devPct * 100).toFixed(2))), borderColor: '#1e8e5a', borderDash: [5, 4], tension: .3, borderWidth: 2, pointRadius: 3 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false },
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } }, tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y?.toFixed(2)} %` } } },
      scales: { y: { min: 0, max: 100, ticks: { callback: (v) => v + ' %' }, grid: { color: '#eef1f5' } }, x: { grid: { display: false } } },
    },
  });
}
