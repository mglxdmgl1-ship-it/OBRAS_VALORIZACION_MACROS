// Centro de reportes: todos los cuadros del informe mensual en PDF y Excel (con fórmulas).
import { S } from '../store.js';
import { esc, nombreMes, mesDe, hoy, fFecha } from '../util.js';
import { icono } from '../ui.js';
import * as R from '../reportes.js';
import { cab, btnExport, enlazarExport, opcionesPeriodo, mesesDeObra } from './comun.js';
import { refrescar } from '../app.js';

let k = null, ym = null, dia = null;

export function render(cont) {
  const m = S.motor;
  const pers = m.periodos();
  if (k === null || k >= pers.length) { const a = pers.findIndex((p) => p.estado !== 'CERRADO'); k = a >= 0 ? a : pers.length - 1; }
  const meses = mesesDeObra();
  if (!ym) ym = pers[k] ? mesDe(pers[k].al) : meses[meses.length - 1] || mesDe(hoy());
  const fs = m.fechasConMetrado();
  if (!dia) dia = fs[fs.length - 1] || hoy();
  const hayV = pers.length > 0 && m.partidasP.length > 0;
  const grupo = (titulo, items) => `<div class="panel"><div class="panel-cab"><h2>${titulo}</h2></div><div class="panel-cuerpo"><div class="rep-lista">${items.map(([id, t, d, ok = true]) => `<div class="rep"><h4>${t}</h4><p>${d}</p><div class="acciones">${ok ? btnExport(id) : '<span class="tenue" style="font-size:12px">Sin datos</span>'}</div></div>`).join('')}</div></div></div>`;
  cont.innerHTML = `
    ${cab('Reportes PDF / Excel', 'Los PDF salen en A4 con el encabezado de la obra y las firmas. Los Excel incluyen <b>fórmulas</b> (parciales, sumas de títulos, acumulados, saldos, kardex, cálculo de insumos) para que pueda seguir trabajándolos.')}
    <div class="panel"><div class="panel-cuerpo"><div class="filtros" style="margin:0">
      <label>Valorización <select id="r-per">${opcionesPeriodo(k) || '<option>—</option>'}</select></label>
      <label>Mes <select id="r-mes">${meses.map((x) => `<option value="${x}" ${x === ym ? 'selected' : ''}>${nombreMes(x)}</option>`).join('')}</select></label>
      <label>Día <input type="date" id="r-dia" value="${dia}"></label>
    </div></div></div>
    ${grupo('Costos y presupuestos', [
      ['pres', 'Presupuesto de obra', 'Títulos y partidas con metrado, P.U. y parcial; resumen con costos indirectos.', m.partidasP.length],
      ['acu', 'Análisis de costos unitarios', 'Todos los A.C.U. con mano de obra, materiales, equipos y subtotales.', m.partidasP.length],
      ['exp', 'Relación de insumos', 'Cantidades totales del presupuesto (metrado × A.C.U.) y su costo.', m.partidasP.length],
      ['ins', 'Catálogo de insumos y precios', 'Precios únicos usados por todos los A.C.U.', S.obra.insumos.length],
    ])}
    ${grupo('Valorización e informe mensual', [
      ['v41', '4.1 Resumen de metrados del mes', `Metrado por partida y por día de ${nombreMes(ym)}.`, m.partidasP.length],
      ['v42', '4.2 Planilla de metrados del día', `Planilla diaria del ${fFecha(dia)} con estados.`, m.partidasP.length],
      ['v43', '4.3 Resumen de valorizaciones', 'Subpresupuestos, costos indirectos y Formato 10.', hayV],
      ['v44', '4.4 Valorización mensual de obra', 'Por partida: presupuesto, anterior, actual, acumulado y saldo.', hayV],
      ['v45', '4.5 Programado vs ejecutado', 'Cronograma comparado del mes y % de cumplimiento.', hayV],
      ['cs', 'Curva S', 'Avance programado, físico y financiero por mes.', m.partidasP.length],
      ['v413', '4.13 Cronograma valorizado', 'Programado mensual por partida.', m.cronoMeses().length],
    ])}
    ${grupo('Almacén y equipo mecánico', [
      ['a410', '4.10 Movimiento diario de almacén', `Entradas y salidas por día de ${nombreMes(ym)}.`, S.obra.movimientos.length],
      ['a411', '4.11 Ingresos – egresos', 'Saldo anterior, ingresos, egresos y saldo del mes.', S.obra.movimientos.length],
      ['a412', '4.12 Saldo valorizado', 'Saldo × P.U. al cierre del mes.', S.obra.movimientos.length],
      ['kar', 'Kardex valorizado (todos)', 'Un kardex por ítem del almacén con fórmulas de saldo.', S.obra.movimientos.length],
      ['bin', 'Bincard (todos)', 'Tarjeta de control visible por ítem.', S.obra.movimientos.length],
      ['mov', 'Movimientos del mes', 'Listado de ingresos y salidas con documentos y vales.', S.obra.movimientos.length],
      ['e415', '4.15 Control de equipo mecánico', 'Horas por día, acumulado, saldo de la O/S y valorizado.', S.obra.equipos.length],
      ['par', 'Partes diarios del mes', 'Horario, horómetro, horas a valorizar, combustible y trabajo.', S.obra.partes.length],
    ])}
    <div class="nota">${icono('alerta')} Las plantillas para importar (presupuesto, A.C.U., metrados, movimientos, precios, cronograma) se descargan desde cada pantalla con el botón «Plantilla».</div>`;
  const ini = ym + '-01';
  enlazarExport(cont, {
    pres: () => R.rPresupuesto(), acu: () => R.rACU(null), exp: () => R.rExplosion(), ins: () => R.rInsumos(),
    v41: () => R.rMetradosMes(ym), v42: () => R.rPlanillaDia(dia), v43: () => R.rResumen(k), v44: () => R.rValorizacion(k), v45: () => R.rComparado(k), cs: () => R.rCurvaS(), v413: () => R.rCronograma(),
    a410: () => R.rMovDiario(ym), a411: () => R.rIngresosEgresos(ym), a412: () => R.rSaldoValorizado(ym), kar: () => R.rKardex(null), bin: () => R.rKardex(null, null, null, true),
    mov: () => { const l = m.movimientos().filter((x) => x.fecha >= ini && x.fecha.startsWith(ym)); return R.rMovimientos(l, 'MOVIMIENTOS DEL ALMACÉN', nombreMes(ym)); },
    e415: () => R.rControlEquipos(ym), par: () => R.rPartes(S.obra.partes.filter((p) => p.fecha.startsWith(ym)).sort((a, b) => (a.fecha < b.fecha ? -1 : 1))),
  });
  cont.querySelector('#r-per').onchange = (e) => { k = Number(e.target.value); ym = mesDe(pers[k].al); refrescar(); };
  cont.querySelector('#r-mes').onchange = (e) => { ym = e.target.value; refrescar(); };
  cont.querySelector('#r-dia').onchange = (e) => { dia = e.target.value; refrescar(); };
  void esc;
}
