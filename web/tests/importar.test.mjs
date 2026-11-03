// Pruebas de los importadores con hojas en formato S10 (filas como arreglos, sin Excel).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { analizarPresupuesto, analizarACU, aplicarACU, importarMetrados, importarCronograma, obraDesdeValorizacionJSON } from '../public/js/importar.js';
import { crearMotor } from '../public/js/calc.js';
import { normalizar } from '../public/js/store.js';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const PRES = [
  ['S10'], ['Presupuesto', '0101001 OBRA DE PRUEBA'], ['Subpresupuesto', '001 ESTRUCTURAS'], [],
  ['Item', 'Descripción', 'Und.', 'Metrado', 'Precio S/.', 'Parcial S/.'],
  ['01', 'OBRAS PROVISIONALES'], ['01.01', 'CARTEL DE OBRA', 'und', 1, 537.01, 537.01], ['01.02', 'TRAZO Y REPLANTEO', 'm2', 100.5, 2.5, 251.25],
  ['', 'Costo Directo', '', '', '', 788.26], ['', 'GASTOS GENERALES (10%)', '', '', '', 78.83],
  ['Subpresupuesto', '002 ARQUITECTURA'], ['01', 'MUROS'], ['01.01', 'MURO DE LADRILLO', 'm2', 50, 80, 4000],
];
const ACU = [
  ['Análisis de precios unitarios'], ['Subpresupuesto', '001 ESTRUCTURAS'],
  ['Partida', '01.01', 'CARTEL DE OBRA'], ['Rendimiento', 'und/DIA', 'MO. 2.0000', 'EQ. 2.0000'],
  ['Código', 'Descripción Recurso', 'Unidad', 'Cuadrilla', 'Cantidad', 'Precio S/.', 'Parcial S/.'], ['', 'Mano de Obra'],
  ['0147010002', 'OPERARIO', 'hh', 1, 4, 18.12, 72.48], ['', 'Materiales'], ['0243010003', 'MADERA TORNILLO', 'p2', '', 30, 8.5, 255],
  ['', 'Equipos'], ['0337010001', 'HERRAMIENTAS MANUALES', '%MO', '', 3, 72.48, 2.17], ['', 'Subpartidas'], ['900101', 'CONCRETO 140', 'm3', '', 0.1, 500, 50],
  ['Subpresupuesto', '002 ARQUITECTURA'], ['Partida', '01.01', 'MURO DE LADRILLO'], ['Rendimiento', 'm2/DIA', 'MO. 10'],
  ['Código', 'Descripción Recurso', 'Unidad', 'Cuadrilla', 'Cantidad', 'Precio S/.', 'Parcial S/.'], ['', 'Materiales'], ['0217010001', 'LADRILLO KK 18 HUECOS', 'und', '', 39, 1.2, 46.8],
  ['Sub Partida', '900101', 'CONCRETO 140'], ['Código', 'Descripción Recurso', 'Unidad', 'Cuadrilla', 'Cantidad', 'Precio S/.', 'Parcial S/.'], ['', 'Materiales'], ['0221000001', 'CEMENTO PORTLAND TIPO I (42.5 kg)', 'bol', '', 8, 28.8, 230.4],
];

test('presupuesto S10 con dos subpresupuestos', () => {
  const r = analizarPresupuesto([{ nombre: 'Hoja1', filas: PRES }]);
  const items = r.partidas.map((p) => p.item);
  assert.deepEqual(items, ['01', '01.01', '01.01.01', '01.01.02', '02', '02.01', '02.01.01']);
  assert.equal(r.partidas.find((p) => p.item === '01.01.02').metrado, 100.5);
  assert.equal(r.avisos.length, 0);
  assert.ok(r.indirectos.some((x) => x.concepto.startsWith('GASTOS GENERALES')));
});

test('A.C.U. S10: secciones, %MO, rendimiento y subpartidas', () => {
  const a = analizarACU([{ nombre: 'ACU', filas: ACU }]);
  assert.deepEqual(Object.keys(a.acu), ['01.01.01', '02.01.01']);
  const f = a.acu['01.01.01'];
  assert.deepEqual(f.map((x) => x.tipo), ['MO', 'MAT', '%MO', 'MAT']);
  assert.equal(f[3].nombre, 'CEMENTO PORTLAND TIPO I (42.5 kg)');
  assert.equal(f[3].cantidad, 0.8);
  assert.equal(a.rendimientos['01.01.01'], 2);
  assert.equal(a.rendimientos['02.01.01'], 10);
  const obra = normalizar({ partidas: analizarPresupuesto([{ nombre: 'P', filas: PRES }]).partidas });
  const r = aplicarACU(obra, a);
  assert.equal(r.partidas, 2);
  assert.equal(r.nuevos, 4);
  const M = crearMotor(obra);
  assert.equal(M.puACU('02.01.01'), 46.8);
  const c = M.consumo([{ item: '02.01.01', metrado: 10 }]);
  assert.equal(c[0].cantidad, 390);
});

test('metrados y cronograma desde tablas', () => {
  const obra = normalizar({ partidas: analizarPresupuesto([{ nombre: 'P', filas: PRES }]).partidas });
  const r = importarMetrados(obra, [['FECHA', 'ITEM', 'METRADO'], ['15/10/2026', '1.1.2', 20], ['16/10/2026', '01.01.02', 5.5], ['16/10/2026', '09.09', 1]]);
  assert.equal(r.n, 2);
  assert.equal(r.errores.length, 1);
  assert.equal(obra.metrados['2026-10-15']['01.01.02'], 20);
  const c = importarCronograma(obra, [['ITEM', 'OCTUBRE 2026', '2026-11'], ['01.01.01', 300, 237.01]]);
  assert.equal(c.n, 1);
  assert.deepEqual(obra.cronograma.meses, ['2026-10', '2026-11']);
  assert.equal(crearMotor(obra).programadoMes('2026-11'), 237.01);
});

test('obra desde el JSON de la valorización (herramientas/exportar_valorizacion_json.py)', () => {
  const val = JSON.parse(readFileSync(fileURLToPath(new URL('../../datos/iei881/valorizacion_01_2026-10.json', import.meta.url)), 'utf8'));
  const o = normalizar(obraDesdeValorizacionJSON(val));
  const M = crearMotor(o);
  assert.equal(M.costoDirecto(), 1063525.99);
  const v = M.valorizacion(0);
  assert.equal(v.cd.act, 54534.96);
  assert.equal(v.total.act, 88811.2);
  assert.equal(M.comparado(0).programado, 60213.3);
});
