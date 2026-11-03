// Pruebas de las fórmulas de docs/04_CALCULO_VALORIZACION.md con los datos de la obra I.E.I. 881.
// Ejecutar:  node --test web/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { crearMotor, periodoSiguiente, cronogramaLineal } from '../public/js/calc.js';
import { r2, redondear, redondearCantidad, num, aISO, cmpItem } from '../public/js/util.js';

const ruta = fileURLToPath(new URL('../public/datos/iei881.json', import.meta.url));
const obra = JSON.parse(readFileSync(ruta, 'utf8'));
const M = crearMotor(obra);
const near = (a, b, tol = 0.005) => assert.ok(Math.abs(a - b) <= tol, `${a} ≠ ${b}`);

test('redondeo igual a Excel', () => {
  assert.equal(r2(164.19 * 6.5), 1067.24);
  assert.equal(r2(18.5 * 610.15), 11287.78);
  assert.equal(r2(-2.345), -2.35);
  assert.equal(redondear(1.0005, 3), 1.001);
  assert.equal(redondearCantidad(3.01, 'arriba'), 4);
  assert.equal(redondearCantidad(3.0000000001, 'arriba'), 3);
  assert.equal(redondearCantidad(3.5, 'normal'), 4);
  assert.equal(redondearCantidad(7.2, 'arriba', 5), 10);
  assert.equal(num('1,234.50'), 1234.5);
  assert.equal(num('1.234,50'), 1234.5);
  assert.equal(num('12,5'), 12.5);
  assert.equal(num('20%'), 0.2);
  assert.equal(aISO('15/10/2026'), '2026-10-15');
  assert.equal(aISO(46310), '2026-10-15');
  assert.ok(cmpItem('01.2', '01.10') < 0);
});

test('presupuesto: 653 partidas, 247 títulos, C.D. y total', () => {
  assert.equal(M.partidasP.length, 653);
  assert.equal(obra.partidas.length - 653, 247);
  assert.equal(M.costoDirecto(), 1063525.99);
  assert.equal(M.presupuestoTotal(), 1731971.91);
  const s = M.subpresupuestos();
  assert.equal(s.length, 8);
  assert.equal(s[0].presupuesto, 127988.54);
  assert.equal(s[1].presupuesto, 393217.1);
});

test('valorización N° 01: C.D. del mes, indirectos y total', () => {
  const v = M.valorizacion(0);
  assert.equal(v.cd.act, 54534.96);
  near(v.cd.pctAct * 100, 5.13);
  const gg = v.ind.find((x) => x.concepto.startsWith('GASTOS GENERALES'));
  const sup = v.ind.find((x) => x.concepto.startsWith('GASTOS DE SUPERVISI'));
  assert.equal(gg.act, 10546.06);
  assert.equal(sup.act, 6667.76);
  assert.equal(v.totalIndirecto.act, 34276.24);
  assert.equal(v.total.act, 88811.2);
  assert.equal(v.total.sal, 1643160.71);
});

test('resumen por subpresupuesto (4.3)', () => {
  const v = M.valorizacion(0);
  assert.equal(v.subp[0].acu, 30095.9);
  near(v.subp[0].pctAcu * 100, 23.51);
  assert.equal(v.subp[1].acu, 24439.06);
  near(v.subp[1].pctAcu * 100, 6.22);
  for (let i = 2; i < 8; i++) assert.equal(v.subp[i].acu, 0);
});

test('partida 01.01.01.01.01 (ejemplo de la documentación)', () => {
  const f = M.valorizacion(0).filas.find((x) => x.item === '01.01.01.01.01');
  assert.equal(f.parcial, 257.14);
  near(f.act.m, 130.53);
  assert.equal(f.act.s, 257.14);
  near(f.pctAcu, 1, 1e-9);
  assert.equal(f.sal.s, 0);
});

test('títulos de la valorización iguales al libro de Excel', () => {
  const ref = JSON.parse(readFileSync(fileURLToPath(new URL('../../datos/iei881/valorizacion_01_2026-10.json', import.meta.url)), 'utf8'));
  const v = M.valorizacion(0);
  const porItem = new Map(v.filas.map((f) => [f.item, f]));
  let n = 0;
  for (const p of ref.valorizacion.partidas) {
    const f = porItem.get(p.item);
    assert.ok(f, p.item);
    near(f.parcial, p.parcial, 0.011);
    near(f.act.s, p.actual.monto || 0, 0.011);
    n++;
  }
  assert.equal(n, 900);
});

test('comparado 4.5 y curva S', () => {
  const c = M.comparado(0);
  assert.equal(c.programado, 60213.3);
  near(c.cumplimiento * 100, 90.57);
  const cs = M.curvaS();
  assert.equal(cs.length, 5);
  const esperado = [5.66, 28.97, 62.79, 95.73, 100];
  cs.forEach((x, i) => near(x.progPct * 100, esperado[i]));
  near(cs[0].ejecPct * 100, 5.13);
  assert.equal(cs[1].ejecutado, null);
});

test('A.C.U.: P.U. y consumo = metrado × cantidad', () => {
  const acu = M.acu('01.01.01.01.01');
  assert.equal(acu.length, 5);
  assert.equal(M.puACU('01.01.01.01.01'), 1.96); // el A.C.U. suma 1.9587 (P.U. del presupuesto 1.97)
  const c = M.consumo([{ item: '01.01.01.01.01', metrado: 100 }]);
  const peon = c.find((x) => x.insumo === 'PEON');
  near(peon.cantidad, 8, 1e-9);
  const yeso = c.find((x) => x.insumo.startsWith('YESO'));
  near(yeso.cantidad, 2.5, 1e-9);
  assert.ok(M.explosion().length > 100);
});

test('estados del control diario (umbral 20 %)', () => {
  // 16/10: limpieza y trazo sin ingreso de madera ni yeso → NO VALORIZAR (como la planilla del día 16)
  const e = M.estadoDia('01.01.01.01.01', '2026-10-16');
  assert.equal(e.estado.cod, 'NO_VALORIZAR');
  const e2 = M.estadoDia('01.01.03.01', '2026-10-16');
  assert.equal(e2.estado.cod, 'VALIDO');
});

test('cierre de periodo: actual → anterior', () => {
  const o2 = JSON.parse(JSON.stringify(obra));
  const M1 = crearMotor(o2);
  const per = o2.periodos[0];
  const cerrado = {};
  for (const p of M1.partidasP) cerrado[p.item] = M1.metradoPeriodo(per, p.item).m;
  per.cerrado = cerrado;
  per.estado = 'CERRADO';
  o2.periodos.push(periodoSiguiente(o2, per));
  const M2 = crearMotor(o2);
  const v2 = M2.valorizacion(1);
  assert.equal(v2.periodo.del, '2026-11-01');
  assert.equal(v2.periodo.al, '2026-11-30');
  assert.equal(v2.cd.ant, 54534.96);
  assert.equal(v2.cd.act, 0);
  assert.equal(v2.total.ant, 88811.2);
  const f = v2.filas.find((x) => x.item === '01.01.01.01.01');
  near(f.ant.m, 130.53);
});

test('almacén: saldos, kardex y cuadros', () => {
  const movs = M.movimientos();
  assert.equal(movs.length, 101);
  const r = M.resumenAlmacen('2026-10');
  const m01 = r.find((x) => x.codigo === 'M-01');
  assert.equal(m01.ingresos, 285);
  assert.equal(m01.salidas, 33 + 28 + 36 + 19);
  assert.equal(m01.saldo, 285 - 116);
  const k = M.kardex('M-01');
  assert.equal(k.filas[k.filas.length - 1].saldo, 169);
});

test('cronograma lineal reparte el parcial completo', () => {
  const c = cronogramaLineal(obra, M);
  assert.equal(c.meses.length, 5);
  const p = M.partidasP[0];
  near(c.montos[p.item].reduce((a, b) => a + b, 0), M.parcial(p), 0.001);
});
