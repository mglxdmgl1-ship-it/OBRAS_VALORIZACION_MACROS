// Definición de todos los reportes (cuadros del informe mensual) para exportar a PDF y Excel.
import { S } from './store.js';
import { r2, fFecha, nombreMes, finDeMes, diasDelMes, mesDe, normal } from './util.js';
import { TIPOS_INSUMO, CONSUMIBLE } from './calc.js';

const M = () => S.motor;
const O = () => S.obra;
const ORD_T = { MO: 0, MAT: 1, EQ: 2, SC: 3, '%MO': 4, SP: 5 };

/** Filas de un árbol (títulos + partidas) con índices de hijos para sumar en Excel. */
function filasArbol(items, valores) {
  const m = M();
  const idx = new Map();
  const filas = items.map((it, i) => {
    idx.set(it, i);
    const p = m.porItem.get(it);
    return { v: valores(p), est: p.tipo === 'T' ? 'titulo' : null, nivel: m.nivel(it), hijos: p.tipo === 'T' ? [] : null, _it: it };
  });
  filas.forEach((f) => {
    let pa = m.padre.get(f._it);
    while (pa && !idx.has(pa)) pa = m.padre.get(pa);
    if (pa) filas[idx.get(pa)].hijos.push(idx.get(f._it));
  });
  const raices = filas.map((f, i) => (m.padre.get(f._it) && idx.has(m.padre.get(f._it)) ? -1 : i)).filter((i) => i >= 0);
  return { filas, raices };
}
/** Ítems a mostrar: las partidas filtradas y todos sus títulos ascendentes. */
function conTitulos(partidas) {
  const m = M();
  const set = new Set();
  for (const it of partidas) {
    set.add(it);
    let pa = m.padre.get(it);
    while (pa) { set.add(pa); pa = m.padre.get(pa); }
  }
  return O().partidas.filter((p) => set.has(p.item)).map((p) => p.item);
}

// =================================================================== PRESUPUESTO
export function rPresupuesto() {
  const m = M();
  const { filas, raices } = filasArbol(O().partidas.map((p) => p.item), (p) => ({
    item: p.item, descripcion: p.descripcion, und: p.und, metrado: p.tipo === 'P' ? Number(p.metrado) || 0 : null, pu: p.tipo === 'P' ? Number(p.pu) || 0 : null, parcial: m.parciales().get(p.item) || 0,
  }));
  filas.push({ v: { descripcion: 'COSTO DIRECTO', parcial: m.costoDirecto() }, est: 'total', hijos: raices });
  const cdFila = filas.length - 1;
  const ind = m.indirectos();
  const res = [{ v: { concepto: 'COSTO DIRECTO', monto: m.costoDirecto(), pct: 1 }, fx: { monto: (r, L, ctx) => ctx.ref(0, cdFila, 'parcial') } }];
  ind.forEach((x) => res.push({ v: { concepto: x.concepto, monto: x.monto, pct: x.ratio } }));
  res.push({ v: { concepto: 'TOTAL COSTO INDIRECTO', monto: m.totalIndirecto(), pct: m.costoDirecto() ? m.totalIndirecto() / m.costoDirecto() : 0 }, est: 'sub', hijos: ind.map((_, i) => i + 1) });
  res.push({ v: { concepto: 'PRESUPUESTO TOTAL', monto: m.presupuestoTotal() }, est: 'total', hijos: [0, ind.length + 1] });
  return {
    archivo: 'PRESUPUESTO', titulo: 'PRESUPUESTO DE OBRA', subtitulo: `Costo directo S/ ${fmtS(m.costoDirecto())} · Presupuesto total S/ ${fmtS(m.presupuestoTotal())}`, orientacion: 'p',
    secciones: [
      { columnas: [{ id: 'item', t: 'ÍTEM', w: 13, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 52, tipo: 'txt', sangria: true }, { id: 'und', t: 'UND', w: 6, tipo: 'txt' }, { id: 'metrado', t: 'METRADO', w: 11, tipo: 'n2' }, { id: 'pu', t: 'P.U. (S/)', w: 11, tipo: 'n2' }, { id: 'parcial', t: 'PARCIAL (S/)', w: 14, tipo: 'n2' }],
        filas, fx: { parcial: (r, L) => `ROUND(${L.metrado}${r}*${L.pu}${r},2)` }, sumar: ['parcial'], editables: ['metrado', 'pu'] },
      { titulo: 'RESUMEN DEL PRESUPUESTO', columnas: [{ id: 'concepto', t: 'CONCEPTO', w: 60, tipo: 'txt' }, { id: 'monto', t: 'MONTO (S/)', w: 18, tipo: 'n2' }, { id: 'pct', t: '% DEL C.D.', w: 12, tipo: 'pct' }],
        filas: res.map((f, i) => (i > 0 && i <= ind.length ? { ...f, fx: { pct: (r, L) => `${L.monto}${r}/${L.monto}${r - i}` } } : f)), sumar: ['monto'], sinCongelar: true },
    ],
  };
}
const fmtS = (v) => new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v || 0);

// =================================================================== A.C.U.
export function rACU(items) {
  const m = M();
  const lista = (items && items.length ? items : m.partidasP.map((p) => p.item)).filter((it) => m.porItem.get(it)?.tipo === 'P');
  const secciones = lista.map((it) => {
    const p = m.porItem.get(it);
    const filasA = m.acu(it).slice().sort((a, b) => (ORD_T[a.tipo] ?? 9) - (ORD_T[b.tipo] ?? 9));
    const filas = [];
    let tipoAnt = null;
    const moFilas = [];
    for (const f of filasA) {
      if (f.tipo !== tipoAnt) { filas.push({ v: { descripcion: (TIPOS_INSUMO[f.tipo] || f.tipo).toUpperCase() }, est: 'sub', sinFx: true }); tipoAnt = f.tipo; }
      const fila = { v: { codigo: f.codigo, descripcion: f.descripcion, und: f.und, cuadrilla: f.cuadrilla ?? null, cantidad: Number(f.cantidad) || 0, precio: f.precio, parcial: f.parcial } };
      if (f.tipo === 'MO') moFilas.push(filas.length);
      if (f.tipo === '%MO') fila.fx = { precio: () => (moFilas.length ? `SUM(${moFilas.map((i) => '{' + i + '}').join(',')})` : '0'), parcial: (r, L) => `${L.cantidad}${r}/100*${L.precio}${r}` };
      filas.push(fila);
    }
    const datos = filas.filter((f) => !f.est).map((f) => filas.indexOf(f));
    filas.push({ v: { descripcion: 'COSTO UNITARIO DIRECTO POR ' + (p.und || 'UND').toUpperCase(), parcial: m.puACU(it) }, est: 'total', hijos: datos });
    return {
      titulo: `Partida ${it}  ${p.descripcion}  ·  Und: ${p.und}  ·  Rendimiento: ${p.rendimiento ?? '—'} ${p.und}/día  ·  P.U. presupuesto: S/ ${fmtS(p.pu)}`,
      hoja: 'A.C.U.',
      columnas: [{ id: 'codigo', t: 'CÓDIGO', w: 10, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN DEL INSUMO', w: 48, tipo: 'txt' }, { id: 'und', t: 'UND', w: 7, tipo: 'txt' }, { id: 'cuadrilla', t: 'CUADRILLA', w: 9, tipo: 'n4' }, { id: 'cantidad', t: 'CANTIDAD', w: 11, tipo: 'n4' }, { id: 'precio', t: 'PRECIO (S/)', w: 11, tipo: 'n2' }, { id: 'parcial', t: 'PARCIAL (S/)', w: 12, tipo: 'n4' }],
      filas, fx: { parcial: (r, L) => `${L.cantidad}${r}*${L.precio}${r}` }, sumar: ['parcial'], sinCongelar: true,
    };
  });
  // las referencias {i} del %MO apuntan a filas de la misma sección: se resuelven al exportar
  secciones.forEach((sec) => sec.filas.forEach((f) => {
    if (f.fx?.precio) {
      const orig = f.fx.precio;
      f.fx.precio = (r, L, ctx) => {
        const fi = sec.filas.indexOf(f);
        return orig().replace(/\{(\d+)\}/g, (_, i) => `${L.parcial}${r - (fi - Number(i))}`);
      };
    }
  }));
  return { archivo: lista.length === 1 ? 'ACU_' + lista[0] : 'ANALISIS_COSTOS_UNITARIOS', titulo: 'ANÁLISIS DE COSTOS UNITARIOS', subtitulo: `${lista.length} partida(s)`, orientacion: 'p', firmas: false, secciones };
}

// =================================================================== INSUMOS
export function rExplosion() {
  const filas = M().explosion().map((x) => ({ v: { codigo: x.codigo, insumo: x.insumo, und: x.und, tipo: x.tipo, cantidad: x.cantidad, precio: x.precio, parcial: r2(x.monto) } }));
  const n = filas.length;
  filas.push({ v: { insumo: 'TOTAL', parcial: r2(filas.reduce((s, f) => s + f.v.parcial, 0)) }, est: 'total', hijos: [...Array(n).keys()] });
  return {
    archivo: 'EXPLOSION_INSUMOS', titulo: 'RELACIÓN DE INSUMOS DEL PRESUPUESTO', subtitulo: 'Cantidades totales = metrado del expediente × cantidad del A.C.U.', orientacion: 'p',
    secciones: [{ columnas: [{ id: 'codigo', t: 'CÓDIGO', w: 10, tipo: 'txt' }, { id: 'insumo', t: 'INSUMO', w: 50, tipo: 'txt' }, { id: 'und', t: 'UND', w: 7, tipo: 'txt' }, { id: 'tipo', t: 'TIPO', w: 6, tipo: 'txt' }, { id: 'cantidad', t: 'CANTIDAD', w: 13, tipo: 'n4' }, { id: 'precio', t: 'PRECIO (S/)', w: 11, tipo: 'n2' }, { id: 'parcial', t: 'PARCIAL (S/)', w: 14, tipo: 'n2' }],
      filas, fx: { parcial: (r, L) => `ROUND(${L.cantidad}${r}*${L.precio}${r},2)` }, sumar: ['parcial'] }],
  };
}

/** Calculadora de insumos: metrados (editables en Excel) → cantidades con fórmulas, redondeo y cantidad final. */
export function rCalculadora(lineas, filasIns, { titulo = 'CÁLCULO DE INSUMOS POR METRADO', subtitulo = '' } = {}) {
  const m = M();
  const fl = lineas.map((l) => {
    const p = m.porItem.get(l.item);
    return { v: { item: l.item, descripcion: p?.descripcion || '', und: p?.und || '', metrado: Number(l.metrado) || 0 } };
  });
  const filas = filasIns.map((x) => ({
    v: { codigo: x.codigo || '', insumo: x.insumo, und: x.und, tipo: x.tipo, teorico: x.cantidad, redondeado: x.redondeado, final: x.final, precio: x.precio, costo: r2(x.final * x.precio), alm: x.alm || '' },
    _det: x.detalle,
  }));
  const nIns = filas.length;
  filas.push({ v: { insumo: 'TOTAL', costo: r2(filas.reduce((s, f) => s + f.v.costo, 0)) }, est: 'total', hijos: [...Array(nIns).keys()] });
  const idxLinea = new Map(lineas.map((l, i) => [l.item, i]));
  const modo = O().parametros?.redondeo || 'arriba';
  const fRed = modo === 'abajo' ? 'ROUNDDOWN' : modo === 'normal' ? 'ROUND' : 'ROUNDUP';
  return {
    archivo: 'CALCULO_INSUMOS', titulo, subtitulo, orientacion: 'l',
    notas: ['Cambie los metrados (celdas amarillas) y el Excel recalcula las cantidades. La columna CANTIDAD FINAL es la que se entrega o se pide (editable).'],
    secciones: [
      { titulo: 'METRADOS', columnas: [{ id: 'item', t: 'ÍTEM', w: 14, tipo: 'txt' }, { id: 'descripcion', t: 'PARTIDA', w: 60, tipo: 'txt' }, { id: 'und', t: 'UND', w: 7, tipo: 'txt' }, { id: 'metrado', t: 'METRADO', w: 12, tipo: 'n2' }], filas: fl, editables: ['metrado'], sinCongelar: true },
      {
        titulo: 'INSUMOS', columnas: [{ id: 'codigo', t: 'CÓDIGO', w: 10, tipo: 'txt' }, { id: 'insumo', t: 'INSUMO', w: 42, tipo: 'txt' }, { id: 'und', t: 'UND', w: 6, tipo: 'txt' }, { id: 'tipo', t: 'TIPO', w: 6, tipo: 'txt' }, { id: 'teorico', t: 'CANTIDAD SEGÚN A.C.U.', w: 13, tipo: 'n4' }, { id: 'redondeado', t: 'REDONDEADO', w: 11, tipo: 'n2' }, { id: 'final', t: 'CANTIDAD FINAL', w: 11, tipo: 'n2' }, { id: 'precio', t: 'PRECIO (S/)', w: 10, tipo: 'n2' }, { id: 'costo', t: 'COSTO (S/)', w: 12, tipo: 'n2' }, { id: 'alm', t: 'EQUIVALENCIA EN ALMACÉN', w: 24, tipo: 'txt' }],
        filas,
        fx: {
          teorico: (r, L, ctx, f) => f._det && f._det.length ? f._det.map((d) => `${ctx.ref(0, idxLinea.get(d.item), 'metrado')}*${+Number(d.unitario).toPrecision(12)}`).join('+') : null,
          redondeado: (r, L) => `${fRed}(${L.teorico}${r},0)`,
          costo: (r, L) => `ROUND(${L.final}${r}*${L.precio}${r},2)`,
        },
        sumar: ['costo'], editables: ['final'], sinCongelar: true,
      },
    ],
  };
}

// =================================================================== METRADOS DIARIOS
export function rPlanillaDia(f) {
  const m = M();
  const items = Object.entries(O().metrados[f] || {}).filter(([, v]) => Number(v) > 0).map(([it]) => it).filter((it) => m.porItem.get(it));
  const lista = conTitulos(items);
  let n = 0;
  const { filas } = filasArbol(lista, (p) => {
    if (p.tipo === 'T') return { n: ++n, item: p.item, descripcion: p.descripcion, parcial: null };
    const e = m.estadoDia(p.item, f);
    const md = m.metradoDia(f, p.item);
    return { n: ++n, item: p.item, descripcion: p.descripcion, und: p.und, metrado: md, acumulado: e?.acumulado ?? md, et: Number(p.metrado) || 0, pct: p.metrado ? (e?.acumulado ?? md) / p.metrado : 0, pu: Number(p.pu) || 0, parcial: r2(md * (Number(p.pu) || 0)), estado: e ? (e.estado.cod === 'VALIDO' ? 'CONFORME' : e.estado.txt + (e.faltantes.length ? ' – falta: ' + e.faltantes.join(', ') : '')) : '' };
  });
  filas.forEach((x) => { if (x.est === 'titulo') { x.v.parcial = null; x.hijos = null; } });
  const tot = r2(items.reduce((s, it) => s + r2(m.metradoDia(f, it) * (Number(m.porItem.get(it).pu) || 0)), 0));
  const datos = filas.map((x, i) => (x.est ? -1 : i)).filter((i) => i >= 0);
  filas.push({ v: { descripcion: 'TOTAL VALORIZADO DEL DÍA (C.D.)', parcial: tot }, est: 'total', hijos: datos });
  const dObra = m.inicio ? Math.round((Date.parse(f) - Date.parse(m.inicio)) / 86400000) + 1 : '';
  const dia = ['DOMINGO', 'LUNES', 'MARTES', 'MIÉRCOLES', 'JUEVES', 'VIERNES', 'SÁBADO'][new Date(f + 'T00:00:00Z').getUTCDay()];
  return {
    archivo: 'PLANILLA_METRADOS_' + f, titulo: 'PLANILLA DE METRADOS DIARIOS', codigo: '4.2', subtitulo: `Fecha: ${fFecha(f)} (${dia}) · Día de obra N° ${dObra} · ${items.length} partida(s)`, orientacion: 'l',
    secciones: [{
      columnas: [{ id: 'n', t: 'N°', w: 4, tipo: 'int' }, { id: 'item', t: 'ÍTEM', w: 13, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 46, tipo: 'txt', sangria: true }, { id: 'und', t: 'UND', w: 5, tipo: 'txt' }, { id: 'metrado', t: 'METRADO DEL DÍA', w: 10, tipo: 'n2' }, { id: 'acumulado', t: 'ACUM. AL DÍA', w: 10, tipo: 'n2' }, { id: 'et', t: 'METRADO E.T.', w: 10, tipo: 'n2' }, { id: 'pct', t: '% ACUM.', w: 8, tipo: 'pct' }, { id: 'pu', t: 'P.U. (S/)', w: 9, tipo: 'n2' }, { id: 'parcial', t: 'PARCIAL DEL DÍA (S/)', w: 11, tipo: 'n2' }, { id: 'estado', t: 'ESTADO / ALERTA', w: 34, tipo: 'txt' }],
      filas, fx: { parcial: (r, L) => `ROUND(${L.metrado}${r}*${L.pu}${r},2)`, pct: (r, L) => `IF(${L.et}${r}=0,0,${L.acumulado}${r}/${L.et}${r})` }, sumar: ['parcial'],
    }],
  };
}

/** Cuadro 4.1: resumen de metrados ejecutados en el mes por partida y por día. */
export function rMetradosMes(ym) {
  const m = M();
  const nd = diasDelMes(ym);
  const ini = ym + '-01', fin = finDeMes(ym);
  const items = m.partidasP.filter((p) => m.sumaDiaria(p.item, ini, fin) > 0).map((p) => p.item);
  const dias = [];
  for (let d = 1; d <= nd; d++) {
    const f = `${ym}-${String(d).padStart(2, '0')}`;
    if (items.some((it) => m.metradoDia(f, it) > 0)) dias.push(f);
  }
  const lista = conTitulos(items);
  const { filas } = filasArbol(lista, (p) => {
    if (p.tipo === 'T') return { item: p.item, descripcion: p.descripcion };
    const v = { item: p.item, descripcion: p.descripcion, und: p.und, et: Number(p.metrado) || 0, pu: Number(p.pu) || 0 };
    v.ant = (Number(O().anterior_inicial?.[p.item]) || 0) + m.sumaDiaria(p.item, null, prevDia(ini));
    let s = 0;
    dias.forEach((f) => { const x = m.metradoDia(f, p.item); v['d' + f] = x || null; s += x; });
    v.mes = s; v.acum = v.ant + s; v.pct = v.et ? v.acum / v.et : 0; v.val = r2(s * v.pu);
    return v;
  });
  filas.forEach((x) => { if (x.est === 'titulo') x.hijos = null; });
  const datos = filas.map((x, i) => (x.est ? -1 : i)).filter((i) => i >= 0);
  filas.push({ v: { descripcion: 'TOTAL VALORIZADO DEL MES (C.D.)', val: r2(datos.reduce((s, i) => s + filas[i].v.val, 0)) }, est: 'total', hijos: datos });
  const colDias = dias.map((f) => ({ id: 'd' + f, t: f.slice(8), w: 6, tipo: 'n2', g: 'METRADO POR DÍA' }));
  return {
    archivo: 'RESUMEN_METRADOS_' + ym, titulo: 'RESUMEN DE METRADOS EJECUTADOS EN EL MES', codigo: '4.1', subtitulo: nombreMes(ym), orientacion: 'l',
    secciones: [{
      columnas: [{ id: 'item', t: 'ÍTEM', w: 12, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 38, tipo: 'txt' }, { id: 'und', t: 'UND', w: 5, tipo: 'txt' }, { id: 'et', t: 'METRADO E.T.', w: 9, tipo: 'n2' }, { id: 'pu', t: 'P.U.', w: 8, tipo: 'n2' }, { id: 'ant', t: 'ACUM. ANTERIOR', w: 9, tipo: 'n2' }, ...colDias, { id: 'mes', t: 'METRADO DEL MES', w: 9, tipo: 'n2' }, { id: 'acum', t: 'ACUMULADO', w: 9, tipo: 'n2' }, { id: 'pct', t: '% AVANCE', w: 8, tipo: 'pct' }, { id: 'val', t: 'VALORIZADO DEL MES (S/)', w: 11, tipo: 'n2' }],
      filas,
      fx: {
        mes: (r, L) => (dias.length ? `SUM(${L['d' + dias[0]]}${r}:${L['d' + dias[dias.length - 1]]}${r})` : '0'),
        acum: (r, L) => `${L.ant}${r}+${L.mes}${r}`, pct: (r, L) => `IF(${L.et}${r}=0,0,${L.acum}${r}/${L.et}${r})`, val: (r, L) => `ROUND(${L.mes}${r}*${L.pu}${r},2)`,
      },
      sumar: ['val'],
    }],
  };
}
function prevDia(f) {
  const d = new Date(f + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() - 1);
  return d.toISOString().slice(0, 10);
}

// =================================================================== VALORIZACIÓN
function subtituloPeriodo(v) {
  const p = v.periodo;
  return `Valorización N° ${String(p.numero).padStart(2, '0')} · Periodo del ${fFecha(p.del)} al ${fFecha(p.al)} · ${nombreMes(mesDe(p.al))}`;
}
export function rValorizacion(k, { soloConAvance = false } = {}) {
  const m = M();
  const v = m.valorizacion(k);
  let filasV = v.filas;
  if (soloConAvance) {
    const items = filasV.filter((f) => f.tipo === 'P' && (f.acu.m || f.act.m)).map((f) => f.item);
    const set = new Set(conTitulos(items));
    filasV = filasV.filter((f) => set.has(f.item));
  }
  const porItem = new Map(filasV.map((f) => [f.item, f]));
  const { filas, raices } = filasArbol(filasV.map((f) => f.item), (p) => {
    const f = porItem.get(p.item);
    return {
      item: f.item, descripcion: f.descripcion, und: f.und, metrado: f.metrado, pu: f.pu, parcial: f.parcial,
      antM: f.ant.m, antS: f.ant.s, antP: f.parcial ? f.ant.s / f.parcial : 0,
      actM: f.act.m, actS: f.act.s, actP: f.pctAct,
      acuM: f.acu.m, acuS: f.acu.s, acuP: f.pctAcu,
      salM: f.sal.m, salS: f.sal.s, salP: f.parcial ? f.sal.s / f.parcial : 0,
    };
  });
  filas.forEach((f) => { if (f.est === 'titulo') ['antP', 'actP', 'acuP', 'salP'].forEach((c) => (f.fx = { ...(f.fx || {}), [c]: (r, L) => `IF(${L.parcial}${r}=0,0,${L[c.replace('P', 'S')]}${r}/${L.parcial}${r})` })); });
  filas.push({ v: { descripcion: 'COSTO DIRECTO', parcial: v.cd.presupuesto, antS: v.cd.ant, actS: v.cd.act, acuS: v.cd.acu, salS: v.cd.sal, antP: v.cd.presupuesto ? v.cd.ant / v.cd.presupuesto : 0, actP: v.cd.pctAct, acuP: v.cd.pctAcu, salP: v.cd.presupuesto ? v.cd.sal / v.cd.presupuesto : 0 }, est: 'total', hijos: raices, fx: Object.fromEntries(['antP', 'actP', 'acuP', 'salP'].map((c) => [c, (r, L) => `IF(${L.parcial}${r}=0,0,${L[c.replace('P', 'S')]}${r}/${L.parcial}${r})`])) });
  const cdIdx = filas.length - 1;
  const pctFx = (c) => (r, L) => `IF(${L.parcial}${r}=0,0,${L[c]}${r}/${L.parcial}${r})`;
  const sec1 = {
    columnas: [
      { id: 'item', t: 'ÍTEM', w: 12, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 40, tipo: 'txt', sangria: true }, { id: 'und', t: 'UND', w: 5, tipo: 'txt' },
      { id: 'metrado', t: 'METRADO', w: 9, tipo: 'n2', g: 'PRESUPUESTO' }, { id: 'pu', t: 'P.U.', w: 8, tipo: 'n2', g: 'PRESUPUESTO' }, { id: 'parcial', t: 'PARCIAL S/', w: 11, tipo: 'n2', g: 'PRESUPUESTO' },
      { id: 'antM', t: 'METRADO', w: 8, tipo: 'n2', g: 'ACUMULADO ANTERIOR' }, { id: 'antS', t: 'S/', w: 10, tipo: 'n2', g: 'ACUMULADO ANTERIOR' }, { id: 'antP', t: '%', w: 6, tipo: 'pct', g: 'ACUMULADO ANTERIOR' },
      { id: 'actM', t: 'METRADO', w: 8, tipo: 'n2', g: 'DEL MES (ACTUAL)' }, { id: 'actS', t: 'S/', w: 10, tipo: 'n2', g: 'DEL MES (ACTUAL)' }, { id: 'actP', t: '%', w: 6, tipo: 'pct', g: 'DEL MES (ACTUAL)' },
      { id: 'acuM', t: 'METRADO', w: 8, tipo: 'n2', g: 'ACUMULADO ACTUAL' }, { id: 'acuS', t: 'S/', w: 10, tipo: 'n2', g: 'ACUMULADO ACTUAL' }, { id: 'acuP', t: '%', w: 6, tipo: 'pct', g: 'ACUMULADO ACTUAL' },
      { id: 'salM', t: 'METRADO', w: 8, tipo: 'n2', g: 'SALDO' }, { id: 'salS', t: 'S/', w: 10, tipo: 'n2', g: 'SALDO' }, { id: 'salP', t: '%', w: 6, tipo: 'pct', g: 'SALDO' },
    ],
    filas,
    fx: {
      parcial: (r, L) => `ROUND(${L.metrado}${r}*${L.pu}${r},2)`, antP: pctFx('antS'),
      actS: (r, L) => `ROUND(${L.actM}${r}*${L.pu}${r},2)`, actP: pctFx('actS'),
      acuM: (r, L) => `${L.antM}${r}+${L.actM}${r}`, acuS: (r, L) => `${L.antS}${r}+${L.actS}${r}`, acuP: pctFx('acuS'),
      salM: (r, L) => `${L.metrado}${r}-${L.acuM}${r}`, salS: (r, L) => `${L.parcial}${r}-${L.acuS}${r}`, salP: pctFx('salS'),
    },
    sumar: ['parcial', 'antS', 'actS', 'acuS', 'salS'], editables: ['actM'],
  };
  const sec2 = seccionResumen(v, (col) => (ctx) => ctx.ref(0, cdIdx, col));
  const idxT = new Map(filas.map((f, i) => [f.v.item, i]));
  sec2.filas.forEach((f) => {
    const i = idxT.get(f.v.item);
    if (f.v.item && i !== undefined) f.fx = { presupuesto: (r, L, ctx) => ctx.ref(0, i, 'parcial'), ant: (r, L, ctx) => ctx.ref(0, i, 'antS'), act: (r, L, ctx) => ctx.ref(0, i, 'actS') };
  });
  return { archivo: 'VALORIZACION_' + String(v.periodo.numero).padStart(2, '0'), titulo: 'VALORIZACIÓN MENSUAL DE OBRA', codigo: '4.4', subtitulo: subtituloPeriodo(v), orientacion: 'l', secciones: [sec1, sec2] };
}

/** Sección del resumen 4.3 (subpresupuestos + costos indirectos). `refCD` enlaza el C.D. con otra sección. */
function seccionResumen(v, refCD = null) {
  const filas = v.subp.map((s) => ({ v: { item: s.item, concepto: s.descripcion, presupuesto: s.presupuesto, ant: s.ant, act: s.act, acu: s.acu, pctAcu: s.pctAcu, sal: s.sal } }));
  const n = filas.length;
  filas.push({ v: { concepto: 'COSTO DIRECTO', presupuesto: v.cd.presupuesto, ant: v.cd.ant, act: v.cd.act, acu: v.cd.acu, pctAcu: v.cd.pctAcu, sal: v.cd.sal }, est: 'sub', hijos: [...Array(n).keys()] });
  const cdI = filas.length - 1;
  v.ind.forEach((x) => filas.push({
    v: { concepto: `${x.concepto} (${(x.ratio * 100).toFixed(2)} % C.D.)`, presupuesto: x.presupuesto, ant: x.ant, act: x.act, acu: x.acu, pctAcu: x.pctAcu, sal: x.sal },
  }));
  // reescribe la fórmula del indirecto con la fila real del C.D.
  filas.forEach((f, i) => {
    if (i > cdI) {
      const offset = i - cdI;
      f.fx = { act: (r, L) => `ROUND(${L.presupuesto}${r}/${L.presupuesto}${r - offset}*${L.act}${r - offset},2)` };
    }
  });
  const nInd = v.ind.length;
  filas.push({ v: { concepto: 'TOTAL COSTO INDIRECTO', presupuesto: v.totalIndirecto.presupuesto, ant: v.totalIndirecto.ant, act: v.totalIndirecto.act, acu: v.totalIndirecto.acu, pctAcu: v.totalIndirecto.pctAcu, sal: v.totalIndirecto.sal }, est: 'sub', hijos: [...Array(nInd).keys()].map((i) => cdI + 1 + i) });
  filas.push({ v: { concepto: 'PRESUPUESTO TOTAL', presupuesto: v.total.presupuesto, ant: v.total.ant, act: v.total.act, acu: v.total.acu, pctAcu: v.total.pctAcu, sal: v.total.sal }, est: 'total', hijos: [cdI, filas.length - 1] });
  if (refCD) {
    filas[cdI].fx = { presupuesto: (r, L, ctx) => refCD('parcial')(ctx), ant: (r, L, ctx) => refCD('antS')(ctx), act: (r, L, ctx) => refCD('actS')(ctx) };
    filas[cdI].hijos = null;
  }
  return {
    titulo: 'RESUMEN DE LA VALORIZACIÓN (4.3)',
    columnas: [{ id: 'item', t: 'ÍTEM', w: 8, tipo: 'txt' }, { id: 'concepto', t: 'CONCEPTO', w: 50, tipo: 'txt' }, { id: 'presupuesto', t: 'PRESUPUESTO (S/)', w: 14, tipo: 'n2' }, { id: 'ant', t: 'ANTERIOR (S/)', w: 13, tipo: 'n2' }, { id: 'act', t: 'DEL MES (S/)', w: 13, tipo: 'n2' }, { id: 'acu', t: 'ACUMULADO (S/)', w: 13, tipo: 'n2' }, { id: 'pctAcu', t: '% ACUM.', w: 8, tipo: 'pct' }, { id: 'sal', t: 'SALDO (S/)', w: 13, tipo: 'n2' }],
    filas,
    fx: { acu: (r, L) => `${L.ant}${r}+${L.act}${r}`, pctAcu: (r, L) => `IF(${L.presupuesto}${r}=0,0,${L.acu}${r}/${L.presupuesto}${r})`, sal: (r, L) => `${L.presupuesto}${r}-${L.acu}${r}` },
    sumar: ['presupuesto', 'ant', 'act'], sinCongelar: true,
    _todasFx: true,
  };
}
export function rResumen(k) {
  const v = M().valorizacion(k);
  const sec = seccionResumen(v);
  // en el resumen solo, las filas con est también usan las fórmulas de acumulado/saldo
  sec.filas.forEach((f) => { if (f.est) f.fx = { ...(f.fx || {}), acu: sec.fx.acu, pctAcu: sec.fx.pctAcu, sal: sec.fx.sal }; });
  const f10 = Object.entries(v.f10).map(([o, x]) => ({ v: { concepto: { ET: 'METRADOS DEL EXPEDIENTE TÉCNICO', MAYOR_METRADO: 'MAYORES METRADOS', PARTIDA_NUEVA: 'PARTIDAS NUEVAS', DEDUCTIVO: 'DEDUCTIVOS (−)' }[o] || o, presupuesto: r2(x.presupuesto), ant: r2(x.ant), act: r2(x.act), acu: r2(x.ant + x.act) } }));
  return {
    archivo: 'RESUMEN_VALORIZACION_' + String(v.periodo.numero).padStart(2, '0'), titulo: 'RESUMEN DE VALORIZACIONES', codigo: '4.3', subtitulo: subtituloPeriodo(v), orientacion: 'l',
    secciones: [sec, { titulo: 'FORMATO N° 10 (4.3.1): E.T. + MAYORES METRADOS + PARTIDAS NUEVAS − DEDUCTIVOS', columnas: [{ id: 'concepto', t: 'CONCEPTO', w: 50, tipo: 'txt' }, { id: 'presupuesto', t: 'PRESUPUESTO (S/)', w: 14, tipo: 'n2' }, { id: 'ant', t: 'ANTERIOR (S/)', w: 13, tipo: 'n2' }, { id: 'act', t: 'DEL MES (S/)', w: 13, tipo: 'n2' }, { id: 'acu', t: 'ACUMULADO (S/)', w: 13, tipo: 'n2' }], filas: f10, fx: { acu: (r, L) => `${L.ant}${r}+${L.act}${r}` }, sinCongelar: true }],
  };
}
export function rComparado(k) {
  const m = M();
  const c = m.comparado(k);
  const v = m.valorizacion(k);
  const items = c.filas.filter((f) => f.tipo === 'P' && (f.programado || f.ejecutado)).map((f) => f.item);
  const set = new Set(conTitulos(items));
  const sel = c.filas.filter((f) => set.has(f.item));
  const por = new Map(sel.map((f) => [f.item, f]));
  const { filas, raices } = filasArbol(sel.map((f) => f.item), (p) => {
    const f = por.get(p.item);
    return { item: f.item, descripcion: f.descripcion, parcial: f.parcial, programado: f.programado, ejecutado: f.ejecutado, diferencia: f.diferencia, cumplimiento: f.cumplimiento ?? null };
  });
  filas.push({ v: { descripcion: 'COSTO DIRECTO', programado: c.programado, ejecutado: c.ejecutado, diferencia: c.diferencia, cumplimiento: c.cumplimiento }, est: 'total', hijos: raices });
  const fxs = { diferencia: (r, L) => `${L.ejecutado}${r}-${L.programado}${r}`, cumplimiento: (r, L) => `IF(${L.programado}${r}=0,"",${L.ejecutado}${r}/${L.programado}${r})` };
  filas.forEach((f) => { if (f.est) f.fx = fxs; });
  return {
    archivo: 'COMPARADO_' + c.ym, titulo: 'CRONOGRAMA COMPARADO: PROGRAMADO VS EJECUTADO', codigo: '4.5', subtitulo: subtituloPeriodo(v) + ` · Cumplimiento ${c.cumplimiento !== null ? (c.cumplimiento * 100).toFixed(2) + ' %' : '—'}`, orientacion: 'l',
    secciones: [{ columnas: [{ id: 'item', t: 'ÍTEM', w: 12, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 50, tipo: 'txt', sangria: true }, { id: 'parcial', t: 'PRESUPUESTO (S/)', w: 13, tipo: 'n2' }, { id: 'programado', t: 'PROGRAMADO DEL MES (S/)', w: 13, tipo: 'n2' }, { id: 'ejecutado', t: 'EJECUTADO DEL MES (S/)', w: 13, tipo: 'n2' }, { id: 'diferencia', t: 'DIFERENCIA (S/)', w: 12, tipo: 'n2' }, { id: 'cumplimiento', t: '% CUMPLIMIENTO', w: 10, tipo: 'pct' }], filas, fx: fxs, sumar: ['parcial', 'programado', 'ejecutado'] }],
  };
}
export function rCurvaS() {
  const m = M();
  const cs = m.curvaS();
  const cd = m.costoDirecto(), tot = m.presupuestoTotal();
  const filas = cs.map((x, i) => ({
    v: { n: x.n, mes: nombreMes(x.ym), programado: x.programado, progAcum: x.progAcum, progPct: x.progPct, ejecutado: x.ejecutado, ejecAcum: x.ejecAcum, ejecPct: x.ejecPct, devengado: x.devengado, devPct: x.devPct },
    fx: {
      progAcum: (r, L) => (i === 0 ? `${L.programado}${r}` : `${L.progAcum}${r - 1}+${L.programado}${r}`), progPct: (r, L) => `${L.progAcum}${r}/${cd}`,
      ...(x.ejecutado !== null ? { ejecAcum: (r, L) => (i === 0 ? `${L.ejecutado}${r}` : `N(${L.ejecAcum}${r - 1})+${L.ejecutado}${r}`), ejecPct: (r, L) => `${L.ejecAcum}${r}/${cd}` } : {}),
      ...(x.devengado !== null ? { devPct: (r, L) => `${L.devengado}${r}/${tot}` } : {}),
    },
  }));
  return {
    archivo: 'CURVA_S', titulo: 'CURVA S: AVANCE PROGRAMADO, FÍSICO Y FINANCIERO', subtitulo: `C.D. S/ ${fmtS(cd)} · Presupuesto total S/ ${fmtS(tot)}`, orientacion: 'l',
    secciones: [{ columnas: [{ id: 'n', t: 'N°', w: 5, tipo: 'int' }, { id: 'mes', t: 'MES', w: 16, tipo: 'txt' }, { id: 'programado', t: 'PROGRAMADO (S/)', w: 14, tipo: 'n2', g: 'PROGRAMADO (C.D.)' }, { id: 'progAcum', t: 'ACUMULADO (S/)', w: 14, tipo: 'n2', g: 'PROGRAMADO (C.D.)' }, { id: 'progPct', t: '% ACUM.', w: 9, tipo: 'pct', g: 'PROGRAMADO (C.D.)' }, { id: 'ejecutado', t: 'EJECUTADO (S/)', w: 14, tipo: 'n2', g: 'FÍSICO EJECUTADO (C.D.)' }, { id: 'ejecAcum', t: 'ACUMULADO (S/)', w: 14, tipo: 'n2', g: 'FÍSICO EJECUTADO (C.D.)' }, { id: 'ejecPct', t: '% ACUM.', w: 9, tipo: 'pct', g: 'FÍSICO EJECUTADO (C.D.)' }, { id: 'devengado', t: 'DEVENGADO ACUM. (S/)', w: 14, tipo: 'n2', g: 'FINANCIERO' }, { id: 'devPct', t: '% DEL TOTAL', w: 9, tipo: 'pct', g: 'FINANCIERO' }], filas }],
  };
}
export function rCronograma() {
  const m = M();
  const meses = m.cronoMeses();
  const { filas, raices } = filasArbol(O().partidas.map((p) => p.item), (p) => {
    const v = { item: p.item, descripcion: p.descripcion, parcial: m.parciales().get(p.item) || 0 };
    let s = 0;
    meses.forEach((ym) => { const x = p.tipo === 'P' ? m.programadoItemMes(p.item, ym) : 0; v['m' + ym] = x; s += x; });
    v.total = r2(s);
    return v;
  });
  const tot = m.sumarArbol((p) => meses.reduce((s, ym) => s + m.programadoItemMes(p.item, ym), 0));
  meses.forEach((ym) => { const t = m.sumarArbol((p) => m.programadoItemMes(p.item, ym)); filas.forEach((f) => { if (f.est === 'titulo') f.v['m' + ym] = t.get(f.v.item) || 0; }); });
  filas.forEach((f) => { if (f.est === 'titulo') f.v.total = tot.get(f.v.item) || 0; });
  const totalRow = { descripcion: 'COSTO DIRECTO', parcial: m.costoDirecto() };
  meses.forEach((ym) => (totalRow['m' + ym] = m.programadoMes(ym)));
  filas.push({ v: totalRow, est: 'total', hijos: raices });
  const colM = meses.map((ym) => ({ id: 'm' + ym, t: nombreMes(ym, true), w: 11, tipo: 'n2', g: 'PROGRAMADO POR MES (S/)' }));
  return {
    archivo: 'CRONOGRAMA_VALORIZADO', titulo: 'CRONOGRAMA VALORIZADO DE OBRA', codigo: '4.13', subtitulo: `${meses.length} mes(es)`, orientacion: 'l',
    secciones: [{ columnas: [{ id: 'item', t: 'ÍTEM', w: 12, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 40, tipo: 'txt', sangria: true }, { id: 'parcial', t: 'PRESUPUESTO (S/)', w: 12, tipo: 'n2' }, ...colM, { id: 'total', t: 'TOTAL (S/)', w: 12, tipo: 'n2' }],
      filas, fx: { total: (r, L) => (meses.length ? `SUM(${L['m' + meses[0]]}${r}:${L['m' + meses[meses.length - 1]]}${r})` : '0') }, sumar: ['parcial', ...meses.map((ym) => 'm' + ym), 'total'], editables: meses.map((ym) => 'm' + ym) }],
  };
}

// =================================================================== ALMACÉN
const colKardex = [{ id: 'fecha', t: 'FECHA', w: 10, tipo: 'fecha' }, { id: 'doc', t: 'DOCUMENTO', w: 24, tipo: 'txt' }, { id: 'proc', t: 'PROCEDENCIA / DESTINO', w: 30, tipo: 'txt' }, { id: 'ent', t: 'CANT.', w: 9, tipo: 'n2', g: 'ENTRADA' }, { id: 'entS', t: 'S/', w: 10, tipo: 'n2', g: 'ENTRADA' }, { id: 'sal', t: 'CANT.', w: 9, tipo: 'n2', g: 'SALIDA' }, { id: 'salS', t: 'S/', w: 10, tipo: 'n2', g: 'SALIDA' }, { id: 'saldo', t: 'CANT.', w: 9, tipo: 'n2', g: 'SALDO' }, { id: 'pu', t: 'P.U.', w: 8, tipo: 'n2', g: 'SALDO' }, { id: 'saldoS', t: 'S/', w: 11, tipo: 'n2', g: 'SALDO' }];
function seccionKardex(cod, desde, hasta, bincard = false) {
  const m = M();
  const k = m.kardex(cod, desde, hasta);
  const a = k.item || { codigo: cod, descripcion: cod, und: '' };
  const filas = [{ v: { doc: 'SALDO INICIAL', saldo: k.saldoInicial, pu: k.pu, saldoS: r2(k.saldoInicial * k.pu) }, est: 'sub', fx: { saldoS: (r, L) => `ROUND(${L.saldo}${r}*${L.pu}${r},2)` } }];
  k.filas.forEach((f) => filas.push({ v: { fecha: f.fecha, doc: [f.tipo, f.documento, f.numero, f.vale && 'VALE ' + f.vale].filter(Boolean).join(' '), proc: f.tipo === 'SALIDA' ? [f.partida && 'Partida ' + f.partida, f.pedido_por].filter(Boolean).join(' · ') : f.procedencia || f.oc_os || '', ent: f.entrada || null, sal: f.salida || null, saldo: f.saldo, pu: k.pu, entS: f.entradaS || null, salS: f.salidaS || null, saldoS: f.saldoS }, est: f.negativo ? 'alerta' : null }));
  const cols = bincard ? colKardex.filter((c) => ['fecha', 'doc', 'proc', 'ent', 'sal', 'saldo'].includes(c.id)).map((c) => ({ ...c, g: undefined })) : colKardex;
  return {
    titulo: `${a.codigo}  ${a.descripcion}  ·  Unidad: ${a.und}${a.insumo_acu ? '  ·  A.C.U.: ' + a.insumo_acu : ''}`, hoja: bincard ? 'BINCARD' : 'KARDEX',
    columnas: cols, filas,
    fx: { saldo: (r, L) => `${L.saldo}${r - 1}+N(${L.ent}${r})-N(${L.sal}${r})`, ...(bincard ? {} : { entS: (r, L) => `IF(${L.ent}${r}="","",ROUND(${L.ent}${r}*${L.pu}${r},2))`, salS: (r, L) => `IF(${L.sal}${r}="","",ROUND(${L.sal}${r}*${L.pu}${r},2))`, saldoS: (r, L) => `ROUND(${L.saldo}${r}*${L.pu}${r},2)` }) },
    sinCongelar: true,
  };
}
export function rKardex(codigos, desde, hasta, bincard = false) {
  const lista = codigos?.length ? codigos : M().almacen.map((a) => a.codigo);
  const secciones = lista.filter((c) => M().movimientos().some((x) => x.codigo === c)).map((c) => seccionKardex(c, desde, hasta, bincard));
  return { archivo: (bincard ? 'BINCARD' : 'KARDEX') + (lista.length === 1 ? '_' + lista[0] : ''), titulo: bincard ? 'TARJETA DE CONTROL VISIBLE DE ALMACÉN (BINCARD)' : 'KARDEX FÍSICO VALORIZADO', subtitulo: desde || hasta ? `Del ${fFecha(desde) || 'inicio'} al ${fFecha(hasta) || 'la fecha'}` : 'Todos los movimientos', orientacion: bincard ? 'p' : 'l', secciones: secciones.length ? secciones : [{ columnas: colKardex, filas: [] }] };
}
export function rMovDiario(ym) {
  const md = M().movimientoDiario(ym);
  const cols = [{ id: 'codigo', t: 'CÓDIGO', w: 8, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 34, tipo: 'txt' }, { id: 'und', t: 'UND', w: 7, tipo: 'txt' }, { id: 'anterior', t: 'SALDO ANTERIOR', w: 9, tipo: 'n2' }];
  md.dias.forEach((f) => { cols.push({ id: 'e' + f, t: 'E', w: 6, tipo: 'n2', g: f.slice(8) + '/' + f.slice(5, 7) }); cols.push({ id: 's' + f, t: 'S', w: 6, tipo: 'n2', g: f.slice(8) + '/' + f.slice(5, 7) }); });
  cols.push({ id: 'ingresos', t: 'TOTAL ENTRADAS', w: 9, tipo: 'n2' }, { id: 'salidas', t: 'TOTAL SALIDAS', w: 9, tipo: 'n2' }, { id: 'saldo', t: 'SALDO FINAL', w: 9, tipo: 'n2' });
  const filas = md.filas.map((r) => {
    const v = { codigo: r.codigo, descripcion: r.descripcion, und: r.und, anterior: r.anterior, ingresos: r.ingresos, salidas: r.salidas, saldo: r.saldo };
    md.dias.forEach((f) => { v['e' + f] = r.porDia[f]?.e || null; v['s' + f] = r.porDia[f]?.s || null; });
    return { v, est: r.saldo < 0 ? 'alerta' : null };
  });
  const eCols = md.dias.map((f) => 'e' + f), sCols = md.dias.map((f) => 's' + f);
  return {
    archivo: 'MOVIMIENTO_DIARIO_ALMACEN_' + ym, titulo: 'CUADRO DE MOVIMIENTO DIARIO DE ALMACÉN', codigo: '4.10', subtitulo: nombreMes(ym), orientacion: 'l',
    secciones: [{ columnas: cols, filas, fx: { ingresos: (r, L) => (eCols.length ? eCols.map((c) => `N(${L[c]}${r})`).join('+') : '0'), salidas: (r, L) => (sCols.length ? sCols.map((c) => `N(${L[c]}${r})`).join('+') : '0'), saldo: (r, L) => `${L.anterior}${r}+${L.ingresos}${r}-${L.salidas}${r}` } }],
  };
}
export function rIngresosEgresos(ym) {
  const filas = M().resumenAlmacen(ym).filter((r) => r.anterior || r.ingresos || r.salidas).map((r) => ({ v: { codigo: r.codigo, descripcion: r.descripcion, und: r.und, anterior: r.anterior, ingresos: r.ingresos, salidas: r.salidas, saldo: r.saldo }, est: r.saldo < 0 ? 'alerta' : null }));
  return {
    archivo: 'INGRESOS_EGRESOS_' + ym, titulo: 'CUADRO RESUMEN DE ALMACÉN: INGRESOS – EGRESOS', codigo: '4.11', subtitulo: nombreMes(ym), orientacion: 'p',
    secciones: [{ columnas: [{ id: 'codigo', t: 'CÓDIGO', w: 9, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 46, tipo: 'txt' }, { id: 'und', t: 'UND', w: 8, tipo: 'txt' }, { id: 'anterior', t: 'SALDO ANTERIOR', w: 11, tipo: 'n2' }, { id: 'ingresos', t: 'INGRESOS DEL MES', w: 11, tipo: 'n2' }, { id: 'salidas', t: 'EGRESOS DEL MES', w: 11, tipo: 'n2' }, { id: 'saldo', t: 'SALDO ACTUAL', w: 11, tipo: 'n2' }], filas, fx: { saldo: (r, L) => `${L.anterior}${r}+${L.ingresos}${r}-${L.salidas}${r}` } }],
  };
}
export function rSaldoValorizado(ym) {
  const filas = M().resumenAlmacen(ym).filter((r) => r.anterior || r.ingresos || r.salidas).map((r) => ({ v: { codigo: r.codigo, descripcion: r.descripcion, und: r.und, saldo: r.saldo, pu: r.pu, valorizado: r.valorizado } }));
  const n = filas.length;
  filas.push({ v: { descripcion: 'TOTAL SALDO VALORIZADO', valorizado: r2(filas.reduce((s, f) => s + f.v.valorizado, 0)) }, est: 'total', hijos: [...Array(n).keys()] });
  return {
    archivo: 'SALDO_VALORIZADO_' + ym, titulo: 'SALDO VALORIZADO DE MATERIALES', codigo: '4.12', subtitulo: 'Al ' + fFecha(finDeMes(ym)), orientacion: 'p',
    secciones: [{ columnas: [{ id: 'codigo', t: 'CÓDIGO', w: 9, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 50, tipo: 'txt' }, { id: 'und', t: 'UND', w: 8, tipo: 'txt' }, { id: 'saldo', t: 'SALDO (CANT.)', w: 12, tipo: 'n2' }, { id: 'pu', t: 'P.U. (S/)', w: 10, tipo: 'n2' }, { id: 'valorizado', t: 'VALORIZADO (S/)', w: 14, tipo: 'n2' }], filas, fx: { valorizado: (r, L) => `ROUND(${L.saldo}${r}*${L.pu}${r},2)` }, sumar: ['valorizado'] }],
  };
}
export function rMovimientos(lista, titulo = 'MOVIMIENTOS DEL ALMACÉN', sub = '') {
  const m = M();
  const filas = lista.map((x) => { const a = m.itemAlmacen(x.codigo); return { v: { fecha: x.fecha, tipo: x.tipo, doc: [x.documento, x.numero].filter(Boolean).join(' '), vale: x.vale || '', partida: x.partida || '', codigo: x.codigo, descripcion: a?.descripcion || '', und: a?.und || '', cantidad: Number(x.cantidad) || 0, pedido: x.pedido_por || '', proc: x.procedencia || '', saldo: x.saldo ?? null }, est: x.negativo ? 'alerta' : null }; });
  return {
    archivo: 'MOVIMIENTOS_ALMACEN', titulo, subtitulo: sub || `${lista.length} movimiento(s)`, orientacion: 'l',
    secciones: [{ columnas: [{ id: 'fecha', t: 'FECHA', w: 10, tipo: 'fecha' }, { id: 'tipo', t: 'TIPO', w: 10, tipo: 'txt' }, { id: 'doc', t: 'DOCUMENTO', w: 22, tipo: 'txt' }, { id: 'vale', t: 'VALE / PECOSA', w: 10, tipo: 'txt' }, { id: 'partida', t: 'PARTIDA', w: 12, tipo: 'txt' }, { id: 'codigo', t: 'CÓDIGO', w: 8, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 34, tipo: 'txt' }, { id: 'und', t: 'UND', w: 7, tipo: 'txt' }, { id: 'cantidad', t: 'CANTIDAD', w: 10, tipo: 'n2' }, { id: 'pedido', t: 'PEDIDO / RECIBIDO POR', w: 18, tipo: 'txt' }, { id: 'proc', t: 'PROCEDENCIA', w: 22, tipo: 'txt' }, { id: 'saldo', t: 'SALDO DESPUÉS', w: 10, tipo: 'n2' }], filas }],
  };
}
/** Vale de salida / PECOSA de un grupo de salidas. */
export function rVale(lista) {
  const m = M();
  const x = lista[0] || {};
  const filas = lista.map((s, i) => { const a = m.itemAlmacen(s.codigo); return { v: { n: i + 1, codigo: s.codigo, descripcion: a?.descripcion || '', und: a?.und || '', cantidad: Number(s.cantidad) || 0, partida: s.partida || '', obs: s.observacion || '' } }; });
  return {
    archivo: 'VALE_SALIDA_' + (x.vale || x.fecha), titulo: 'VALE DE SALIDA DE ALMACÉN', subtitulo: `N° ${x.vale || '______'} · Fecha ${fFecha(x.fecha)} · Pedido por: ${x.pedido_por || '________________'}`, orientacion: 'p',
    secciones: [{ columnas: [{ id: 'n', t: 'N°', w: 5, tipo: 'int' }, { id: 'codigo', t: 'CÓDIGO', w: 9, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 46, tipo: 'txt' }, { id: 'und', t: 'UND', w: 8, tipo: 'txt' }, { id: 'cantidad', t: 'CANTIDAD', w: 11, tipo: 'n2' }, { id: 'partida', t: 'PARTIDA / DESTINO', w: 14, tipo: 'txt' }, { id: 'obs', t: 'OBSERVACIÓN', w: 22, tipo: 'txt' }], filas }],
  };
}

// =================================================================== EQUIPOS
export function rControlEquipos(ym) {
  const m = M();
  const nd = diasDelMes(ym);
  const datos = m.controlEquipos(ym);
  const cols = [{ id: 'codigo', t: 'CÓDIGO', w: 7, tipo: 'txt' }, { id: 'descripcion', t: 'EQUIPO', w: 26, tipo: 'txt' }, { id: 'unidad', t: 'UND', w: 6, tipo: 'txt' }, { id: 'cantidad', t: 'CONTRA-TADO', w: 8, tipo: 'n2' }, { id: 'anterior', t: 'ACUM. ANTERIOR', w: 8, tipo: 'n2' }];
  for (let d = 1; d <= nd; d++) cols.push({ id: 'd' + d, t: String(d), w: 4.2, tipo: 'n2', g: 'HORAS POR DÍA' });
  cols.push({ id: 'mes', t: 'DEL MES', w: 8, tipo: 'n2' }, { id: 'acumulado', t: 'ACUMU-LADO', w: 8, tipo: 'n2' }, { id: 'saldo', t: 'SALDO O/S', w: 8, tipo: 'n2' }, { id: 'costo', t: 'COSTO UNIT.', w: 8, tipo: 'n2' }, { id: 'valorizado', t: 'VALORIZADO MES (S/)', w: 10, tipo: 'n2' });
  const filas = datos.map((e) => { const v = { codigo: e.codigo, descripcion: e.descripcion, unidad: e.unidad, cantidad: e.cantidad, anterior: e.anterior, mes: e.mes, acumulado: e.acumulado, saldo: e.saldo, costo: e.costo, valorizado: e.valorizado }; e.dias.forEach((h, i) => (v['d' + (i + 1)] = h || null)); return { v, est: e.excede ? 'alerta' : null }; });
  const n = filas.length;
  filas.push({ v: { descripcion: 'TOTAL', valorizado: r2(datos.reduce((s, e) => s + e.valorizado, 0)) }, est: 'total', hijos: [...Array(n).keys()] });
  return {
    archivo: 'CONTROL_EQUIPO_MECANICO_' + ym, titulo: 'CONTROL DE EQUIPO MECÁNICO', codigo: '4.15', subtitulo: nombreMes(ym), orientacion: 'l',
    secciones: [{ columnas: cols, filas, fx: { mes: (r, L) => `SUM(${L.d1}${r}:${L['d' + nd]}${r})`, acumulado: (r, L) => `${L.anterior}${r}+${L.mes}${r}`, saldo: (r, L) => `${L.cantidad}${r}-${L.acumulado}${r}`, valorizado: (r, L) => `ROUND(${L.mes}${r}*${L.costo}${r},2)` }, sumar: ['valorizado'] }],
  };
}
export function rPartes(lista) {
  const m = M();
  const filas = lista.map((p) => { const h = m.horasParte(p); const e = (O().equipos || []).find((x) => x.codigo === p.equipo); return { v: { fecha: p.fecha, n: p.n_parte || '', equipo: p.equipo + ' ' + (e?.descripcion || ''), operador: p.operador || '', horario: [p.m_de && `${p.m_de}-${p.m_a}`, p.t_de && `${p.t_de}-${p.t_a}`].filter(Boolean).join(' / '), h_ini: p.h_ini ?? null, h_fin: p.h_fin ?? null, horas: h.valorizar, gal: Number(p.combustible) || null, partida: p.partida || '', trabajo: p.trabajo || '', metrado: Number(p.metrado) || null }, est: h.alertas.length ? 'alerta' : null }; });
  const n = filas.length;
  filas.push({ v: { equipo: 'TOTAL', horas: r2(filas.reduce((s, f) => s + (f.v.horas || 0), 0)), gal: r2(filas.reduce((s, f) => s + (f.v.gal || 0), 0)) }, est: 'total', hijos: [...Array(n).keys()] });
  return {
    archivo: 'PARTES_DIARIOS_EQUIPO', titulo: 'PARTES DIARIOS DE MAQUINARIA Y EQUIPO', subtitulo: `${lista.length} parte(s)`, orientacion: 'l',
    secciones: [{ columnas: [{ id: 'fecha', t: 'FECHA', w: 10, tipo: 'fecha' }, { id: 'n', t: 'N° PARTE', w: 8, tipo: 'txt' }, { id: 'equipo', t: 'EQUIPO', w: 26, tipo: 'txt' }, { id: 'operador', t: 'OPERADOR', w: 18, tipo: 'txt' }, { id: 'horario', t: 'HORARIO', w: 18, tipo: 'txt' }, { id: 'h_ini', t: 'HORÓM. INICIAL', w: 9, tipo: 'n2' }, { id: 'h_fin', t: 'HORÓM. FINAL', w: 9, tipo: 'n2' }, { id: 'horas', t: 'HORAS A VALORIZAR', w: 9, tipo: 'n2' }, { id: 'gal', t: 'COMBUST. (GAL)', w: 8, tipo: 'n2' }, { id: 'partida', t: 'PARTIDA', w: 11, tipo: 'txt' }, { id: 'trabajo', t: 'TRABAJO REALIZADO', w: 34, tipo: 'txt' }, { id: 'metrado', t: 'METRADO', w: 8, tipo: 'n2' }], filas, sumar: ['horas', 'gal'] }],
  };
}

// =================================================================== CATÁLOGOS
export function rInsumos() {
  const filas = (O().insumos || []).map((i) => ({ v: { codigo: i.codigo, descripcion: i.descripcion, und: i.und, tipo: i.tipo, precio: Number(i.precio) || 0, controlar: i.controlar === false ? 'NO' : 'SI' } }));
  return { archivo: 'CATALOGO_INSUMOS', titulo: 'CATÁLOGO DE INSUMOS Y PRECIOS', subtitulo: `${filas.length} insumos`, orientacion: 'p', firmas: false, secciones: [{ columnas: [{ id: 'codigo', t: 'CÓDIGO', w: 10, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 55, tipo: 'txt' }, { id: 'und', t: 'UND', w: 8, tipo: 'txt' }, { id: 'tipo', t: 'TIPO', w: 7, tipo: 'txt' }, { id: 'precio', t: 'PRECIO (S/)', w: 12, tipo: 'n2' }, { id: 'controlar', t: 'CONTROLAR', w: 10, tipo: 'txt' }], filas, editables: ['precio'] }] };
}
export function rCatalogoAlmacen() {
  const st = M().stock();
  const filas = M().almacen.map((a) => ({ v: { codigo: a.codigo, descripcion: a.descripcion, und: a.und, grupo: a.grupo, pu: Number(a.pu) || 0, acu: a.insumo_acu || '', factor: Number(a.factor) || 1, saldo: st.get(a.codigo)?.saldo ?? 0 } }));
  return { archivo: 'CATALOGO_ALMACEN', titulo: 'CATÁLOGO DEL ALMACÉN', subtitulo: `${filas.length} ítems`, orientacion: 'l', firmas: false, secciones: [{ columnas: [{ id: 'codigo', t: 'CÓDIGO', w: 9, tipo: 'txt' }, { id: 'descripcion', t: 'DESCRIPCIÓN', w: 40, tipo: 'txt' }, { id: 'und', t: 'UND', w: 8, tipo: 'txt' }, { id: 'grupo', t: 'GRUPO', w: 14, tipo: 'txt' }, { id: 'pu', t: 'P.U. (S/)', w: 10, tipo: 'n2' }, { id: 'acu', t: 'INSUMO DEL A.C.U.', w: 40, tipo: 'txt' }, { id: 'factor', t: 'FACTOR (und A.C.U. por und almacén)', w: 12, tipo: 'n4' }, { id: 'saldo', t: 'SALDO', w: 10, tipo: 'n2' }], filas }] };
}

/** Plantillas vacías para importar desde Excel. */
export function rPlantilla(tipo) {
  const P = {
    presupuesto: { titulo: 'PLANTILLA: PRESUPUESTO', cols: [['ITEM', 14], ['DESCRIPCION', 60], ['UND', 8], ['METRADO', 12], ['PRECIO', 12], ['PARCIAL', 14]], ej: [['01', 'OBRAS PROVISIONALES', '', '', '', ''], ['01.01', 'CARTEL DE OBRA', 'und', 1, 537.01, 537.01]] },
    acu: { titulo: 'PLANTILLA: A.C.U. (una fila por insumo)', cols: [['PARTIDA', 14], ['INSUMO', 50], ['UND', 8], ['TIPO', 8], ['CUADRILLA', 10], ['CANTIDAD', 12], ['PRECIO', 12]], ej: [['01.01', 'OPERARIO', 'hh', 'MO', 1, 0.8, 18.12], ['01.01', 'CEMENTO PORTLAND TIPO IP (42.5 KG)', 'bol', 'MAT', '', 0.41, 28.8], ['01.01', 'HERRAMIENTAS MANUALES', '%mo', '%MO', '', 3, '']] },
    metrados: { titulo: 'PLANTILLA: METRADOS DIARIOS', cols: [['FECHA', 12], ['ITEM', 14], ['METRADO', 12]], ej: [['15/10/2026', '01.01', 25.5]] },
    movimientos: { titulo: 'PLANTILLA: MOVIMIENTOS DE ALMACÉN', cols: [['FECHA', 12], ['TIPO', 12], ['CODIGO', 10], ['DESCRIPCION', 40], ['UND', 8], ['CANTIDAD', 12], ['PU', 10], ['DOCUMENTO', 18], ['NUMERO', 14], ['PROCEDENCIA', 30], ['PARTIDA', 14], ['PEDIDO_POR', 20], ['VALE', 10], ['ESTADO', 14]], ej: [['16/10/2026', 'INGRESO', 'M-01', 'CEMENTO PORTLAND IP 42.5 KG', 'BOLSA', 300, 28.8, 'GUÍA DE REMISIÓN', '001-123', 'PROVEEDOR S.A.C.', '', '', '', 'CONFIRMADO'], ['17/10/2026', 'SALIDA', 'M-01', '', '', 12, '', 'VALE', '', '', '02.01.01', 'MAESTRO DE OBRA', '001', '']] },
    insumos: { titulo: 'PLANTILLA: PRECIOS DE INSUMOS', cols: [['DESCRIPCION', 50], ['UND', 8], ['TIPO', 8], ['PRECIO', 12]], ej: [['CEMENTO PORTLAND TIPO IP O HS ( 42.5 KG )', 'bol', 'MAT', 28.8]] },
    cronograma: { titulo: 'PLANTILLA: CRONOGRAMA VALORIZADO (S/ por mes)', cols: [['ITEM', 14], ['2026-10', 12], ['2026-11', 12], ['2026-12', 12]], ej: [['01.01', 500, 37.01, 0]] },
  }[tipo];
  return { archivo: 'PLANTILLA_' + tipo.toUpperCase(), titulo: P.titulo, orientacion: 'l', firmas: false, secciones: [{ columnas: P.cols.map(([t, w]) => ({ id: t, t, w, tipo: 'txt' })), filas: P.ej.map((r) => ({ v: Object.fromEntries(P.cols.map(([t], i) => [t, r[i]])) })) }] };
}

export { normal, CONSUMIBLE };
