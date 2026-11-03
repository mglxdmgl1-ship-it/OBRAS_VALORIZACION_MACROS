// Importación desde Excel (S10, Delphin, Kennedy u otros), CSV y JSON.
// Sigue la lógica de ModImportar.bas: detecta encabezados de columnas y de partida
// («Partida 01.01.01» o «01.01.01 DESCRIPCIÓN»), rendimiento, secciones MO/MAT/EQ/SC/subpartidas.
import { num, normal, normItem, aISO, esNumero, r2, uid, mesDe, MESES } from './util.js';

// ------------------------------------------------------------------ lectura de archivos
function valorCelda(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'object') {
    if (v instanceof Date) return v;
    if ('result' in v) return valorCelda(v.result);
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if ('text' in v) return v.text;
    if (v.error) return null;
    return null;
  }
  return v;
}
/** Lee un archivo (.xlsx, .xlsm o .csv) y devuelve [{nombre, filas: any[][]}]. */
export async function leerLibro(file) {
  const nombre = file.name.toLowerCase();
  if (nombre.endsWith('.csv') || nombre.endsWith('.txt')) {
    const txt = await file.text();
    return [{ nombre: file.name, filas: parseCSV(txt) }];
  }
  if (nombre.endsWith('.xls')) throw new Error('El formato .xls (Excel 97-2003) no se puede leer en el navegador. Ábralo en Excel y guárdelo como .xlsx.');
  const wb = new window.ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  return wb.worksheets.map((ws) => {
    const filas = [];
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      const fila = [];
      row.eachCell({ includeEmpty: true }, (c, ci) => { fila[ci - 1] = valorCelda(c.value); });
      filas[r - 1] = fila;
    });
    for (let i = 0; i < filas.length; i++) if (!filas[i]) filas[i] = [];
    return { nombre: ws.name, filas };
  });
}
function parseCSV(txt) {
  const sep = (txt.split('\n')[0].match(/;/g) || []).length > (txt.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
  const out = [];
  let fila = [], campo = '', q = false;
  for (let i = 0; i < txt.length; i++) {
    const c = txt[i];
    if (q) {
      if (c === '"' && txt[i + 1] === '"') { campo += '"'; i++; } else if (c === '"') q = false; else campo += c;
    } else if (c === '"') q = true;
    else if (c === sep) { fila.push(campo); campo = ''; }
    else if (c === '\n') { fila.push(campo); out.push(fila); fila = []; campo = ''; }
    else if (c !== '\r') campo += c;
  }
  if (campo || fila.length) { fila.push(campo); out.push(fila); }
  return out.map((f) => f.map((x) => (esNumero(x) && !/^0\d/.test(x.trim()) ? num(x) : x)));
}
const T = (v) => (v === null || v === undefined ? '' : v instanceof Date ? v.toISOString().slice(0, 10) : String(v).trim());
const H = (v) => normal(T(v)).replace(/[^A-Z0-9%]/g, '');
const RE_ITEM = /^\d{1,4}(\.\d{1,4})*$/;

function textoItem(v) {
  if (typeof v === 'number') {
    if (Number.isInteger(v)) return String(v).padStart(2, '0');
    return String(v);
  }
  const s = normItem(T(v));
  return RE_ITEM.test(s) ? s : '';
}
/** Clave numérica de un ítem para comparar "1.1.1" con "01.01.01". */
export function claveItem(it) {
  return String(it).split('.').map((x) => String(parseInt(x, 10))).join('.');
}

// ------------------------------------------------------------------ PRESUPUESTO
const ENC_PRES = {
  item: (t) => ['ITEM', 'ITEMS', 'COD', 'CODIGO', 'NITEM', 'PARTIDAN', 'ITEMN'].includes(t),
  desc: (t) => t.startsWith('DESCRIP') || t === 'PARTIDA' || t === 'PARTIDAS' || t === 'NOMBREDELAPARTIDA',
  und: (t) => ['UND', 'UNID', 'UNIDAD', 'UM', 'UNIDADDEMEDIDA', 'U', 'UN'].includes(t),
  met: (t) => t.startsWith('METRADO') || t.startsWith('CANT'),
  pu: (t) => t.startsWith('PRECIO') || ['PU', 'CU', 'PUNITARIO', 'COSTO', 'PUS'].includes(t) || t.startsWith('PUNI') || t.startsWith('COSTOUNIT'),
  par: (t) => t.startsWith('PARCIAL') || t === 'TOTAL' || t === 'SUBTOTAL',
};
function buscarEncabezado(filas, defs, minimo) {
  for (let r = 0; r < Math.min(filas.length, 60); r++) {
    const f = filas[r] || [];
    const map = {};
    f.forEach((v, c) => {
      const t = H(v);
      if (!t) return;
      for (const [k, fn] of Object.entries(defs)) if (map[k] === undefined && fn(t) && !Object.values(map).includes(c)) { map[k] = c; break; }
    });
    if (Object.keys(map).length >= minimo) return { fila: r, map };
  }
  return null;
}
/** Analiza un libro de presupuesto. Devuelve {partidas, indirectos, avisos, hoja}. */
export function analizarPresupuesto(hojas) {
  let mejor = null;
  for (const h of hojas) {
    const r = presupuestoDeHoja(h.filas);
    if (!mejor || r.partidas.filter((p) => p.tipo === 'P').length > mejor.partidas.filter((p) => p.tipo === 'P').length) mejor = { ...r, hoja: h.nombre };
  }
  if (!mejor || !mejor.partidas.length) throw new Error('No se encontraron partidas. El archivo debe tener columnas ITEM, DESCRIPCIÓN, UND, METRADO, PRECIO y PARCIAL, con ítems como 01.01.01.');
  return mejor;
}
function presupuestoDeHoja(filas) {
  const avisos = [];
  const enc = buscarEncabezado(filas, ENC_PRES, 3);
  let map, desde;
  if (enc && enc.map.item !== undefined && enc.map.desc !== undefined) { map = enc.map; desde = enc.fila + 1; }
  else { map = { item: 0, desc: 1, und: 2, met: 3, pu: 4, par: 5 }; desde = 0; avisos.push('No se encontró la fila de títulos: se asumieron las columnas A a F.'); }
  const partidas = [];
  const vistos = new Set();
  const indirectos = [];
  let prefijo = '';
  const RE_SUB = /^SUB\s*PRESUPUESTO\s*:?\s*(\d+)\s*(.*)$/i;
  const subs = new Set();
  filas.forEach((f) => { const m = (f || []).map(T).filter(Boolean).join(' ').match(RE_SUB); if (m) subs.add(parseInt(m[1], 10)); });
  const haySub = subs.size >= 2;
  for (let r = 0; r < filas.length; r++) {
    const f = filas[r] || [];
    const txt = f.map(T).filter(Boolean).join(' ');
    const ms = txt.match(RE_SUB);
    if (ms && !haySub) continue;
    if (!ms && r < desde) continue;
    if (ms) {
      prefijo = String(parseInt(ms[1], 10)).padStart(2, '0');
      if (!vistos.has(prefijo)) { partidas.push({ item: prefijo, tipo: 'T', descripcion: (ms[2] || 'SUBPRESUPUESTO ' + prefijo).trim().toUpperCase(), und: '', metrado: null, pu: null }); vistos.add(prefijo); }
      continue;
    }
    let item = textoItem(f[map.item]);
    let desc = T(f[map.desc]);
    if (!item && map.item === map.desc - 0) { /* nada */ }
    if (!item) {
      // «01.01.02 DESCRIPCIÓN» en una sola celda
      const m = T(f[map.item] ?? f[map.desc]).match(/^(\d{1,4}(?:\.\d{1,4})+)\s+(.+)$/);
      if (m) { item = m[1]; desc = m[2]; }
    }
    if (!item) {
      const d = normal(txt);
      const n = f.filter((x) => typeof x === 'number');
      if (n.length && /(GASTOS|UTILIDAD|SUPERVISI|LIQUIDACI|EXPEDIENTE|MONITOREO|SEGUIMIENTO|COSTO DIRECTO|I\.?G\.?V|TOTAL)/.test(d)) indirectos.push({ concepto: T(f.find((x) => typeof x === 'string' && x.trim().length > 3)) || d, monto: n.reduce((mx, x) => (Math.abs(x) > Math.abs(mx) ? x : mx), 0) });
      continue;
    }
    if (haySub && prefijo) item = prefijo + '.' + item;
    if (!desc) continue;
    const und = T(f[map.und]);
    const met = map.met !== undefined ? num(f[map.met]) : 0;
    const pu = map.pu !== undefined ? num(f[map.pu]) : 0;
    const par = map.par !== undefined ? num(f[map.par]) : 0;
    const esP = !!und && (met !== 0 || pu !== 0 || par !== 0);
    if (vistos.has(item)) { avisos.push(`Ítem repetido ${item} (fila ${r + 1}): se omitió.`); continue; }
    vistos.add(item);
    const p = { item, tipo: esP ? 'P' : 'T', descripcion: desc.replace(/\s+/g, ' ').trim(), und: esP ? und : '', metrado: esP ? met : null, pu: esP ? pu : null, rendimiento: null, origen: 'ET' };
    if (esP && !pu && met && par) p.pu = r2(par / met);
    if (esP && par && Math.abs(r2(p.metrado * p.pu) - par) > 0.05) avisos.push(`${item}: metrado × P.U. = ${r2(p.metrado * p.pu)} y el archivo dice ${par}.`);
    partidas.push(p);
  }
  // títulos implícitos que falten (p. ej. 01.02 sin fila propia)
  const set = new Set(partidas.map((p) => p.item));
  const faltan = [];
  for (const p of partidas) {
    const seg = p.item.split('.');
    for (let n = 1; n < seg.length; n++) {
      const c = seg.slice(0, n).join('.');
      if (!set.has(c)) { set.add(c); faltan.push(c); }
    }
  }
  if (faltan.length) avisos.push(`Se agregaron ${faltan.length} título(s) que no tenían fila: ${faltan.slice(0, 8).join(', ')}${faltan.length > 8 ? '…' : ''}`);
  faltan.forEach((it) => partidas.push({ item: it, tipo: 'T', descripcion: 'TÍTULO ' + it, und: '', metrado: null, pu: null, origen: 'ET' }));
  // Mantener el orden del archivo, ubicando cada título faltante antes de su primer hijo
  const orden = [];
  const pos = new Map();
  partidas.forEach((p, i) => pos.set(p.item, i));
  const primero = (it) => { let m = Infinity; for (const p of partidas) if (p.item === it || p.item.startsWith(it + '.')) m = Math.min(m, pos.get(p.item)); return m; };
  partidas.forEach((p) => orden.push({ p, k: faltan.includes(p.item) ? primero(p.item) - 0.5 : pos.get(p.item) }));
  orden.sort((a, b) => a.k - b.k);
  return { partidas: orden.map((x) => x.p), indirectos, avisos };
}

// ------------------------------------------------------------------ A.C.U.
const ETIQ = (s) => {
  const t = normal(s).replace(/[^A-Z ]/g, '').trim();
  if (['MANO DE OBRA', 'MANO OBRA', 'MO'].includes(t)) return 'MO';
  if (['MATERIALES', 'MATERIAL'].includes(t)) return 'MAT';
  if (['EQUIPOS', 'EQUIPO', 'EQUIPOS Y HERRAMIENTAS', 'HERRAMIENTAS Y EQUIPOS', 'MAQUINARIA Y EQUIPOS', 'EQUIPO Y HERRAMIENTAS', 'MAQUINARIAS', 'MAQUINARIA'].includes(t)) return 'EQ';
  if (['SUBCONTRATOS', 'SUBCONTRATO', 'SUB CONTRATOS', 'SERVICIOS'].includes(t)) return 'SC';
  if (['SUBPARTIDAS', 'SUB PARTIDAS', 'SUB PARTIDA', 'SUBPARTIDA'].includes(t)) return 'SP';
  return null;
};
const ENC_ACU_PLANO = {
  partida: (t) => t === 'PARTIDA' || t === 'ITEM' || t === 'CODIGOPARTIDA' || t === 'ITEMPARTIDA',
  insumo: (t) => t.includes('INSUMO') || t.includes('RECURSO') || t === 'DESCRIPCION',
  und: (t) => ['UND', 'UNID', 'UNIDAD', 'UM'].includes(t),
  tipo: (t) => t === 'TIPO' || t === 'CLASE',
  cuad: (t) => t.startsWith('CUAD'),
  cant: (t) => t.startsWith('CANT'),
  precio: (t) => t.startsWith('PRECIO') || t === 'PU' || t === 'COSTO',
};
function quitarCodigo(s) {
  return T(s).replace(/^\s*(\d{2}\s?\d{5,}|\d{6,}|[A-Z]{1,3}-?\d{2,})\s+/, '').trim();
}
function tipoPorUnidad(und, tipo) {
  const u = normal(und).replace(/[^A-Z%]/g, '');
  if (u.startsWith('%MO') || u === '%') return '%MO';
  if (tipo) return tipo;
  if (u === 'HH') return 'MO';
  if (u === 'HM' || u === 'HE') return 'EQ';
  return 'MAT';
}
/** Analiza un libro de A.C.U. Devuelve {acu: {item: filas[]}, insumos: Map, rendimientos, avisos}. */
export function analizarACU(hojas) {
  let mejor = null;
  for (const h of hojas) {
    const r = acuDeHoja(h.filas);
    if (!mejor || Object.keys(r.acu).length > Object.keys(mejor.acu).length) mejor = { ...r, hoja: h.nombre };
  }
  if (!mejor || !Object.keys(mejor.acu).length) throw new Error('No se encontraron análisis de costos unitarios. Se esperan bloques «Partida 01.01.01 …» con sus insumos, o una tabla con columnas PARTIDA, INSUMO, UND, TIPO, CANTIDAD y PRECIO.');
  return mejor;
}
function acuDeHoja(filas) {
  const avisos = [];
  // ---- formato plano (una fila por insumo)
  const plano = buscarEncabezado(filas, ENC_ACU_PLANO, 4);
  if (plano && plano.map.partida !== undefined && plano.map.insumo !== undefined && plano.map.cant !== undefined) {
    const m = plano.map;
    const acu = {};
    for (let r = plano.fila + 1; r < filas.length; r++) {
      const f = filas[r] || [];
      const it = textoItem(f[m.partida]);
      const nom = T(f[m.insumo]);
      if (!it || !nom) continue;
      const und = T(f[m.und]);
      let tipo = normal(T(f[m.tipo])).replace(/[^A-Z%]/g, '');
      tipo = { MANODEOBRA: 'MO', MATERIAL: 'MAT', MATERIALES: 'MAT', EQUIPO: 'EQ', EQUIPOS: 'EQ', SUBCONTRATO: 'SC' }[tipo] || tipo;
      if (!['MO', 'MAT', 'EQ', 'SC', '%MO', 'SP'].includes(tipo)) tipo = tipoPorUnidad(und, null);
      tipo = tipoPorUnidad(und, tipo);
      (acu[it] = acu[it] || []).push({ tipo, nombre: quitarCodigo(nom), und, cuadrilla: m.cuad !== undefined && T(f[m.cuad]) !== '' ? num(f[m.cuad]) : null, cantidad: num(f[m.cant]), precio: m.precio !== undefined ? num(f[m.precio]) : 0 });
    }
    return { acu, rendimientos: {}, avisos };
  }
  // ---- formato por bloques (S10, Delphin…)
  const acu = {}, rend = {}, subs = {};
  let cur = null, tipo = null, cols = null, esSub = false;
  const filaTxt = (f) => f.map(T).filter(Boolean);
  const RE_SUBP = /^SUB\s*PRESUPUESTO\s*:?\s*(\d+)/i;
  const nSubp = new Set();
  filas.forEach((f) => { const m = filaTxt(f || []).join(' ').match(RE_SUBP); if (m) nSubp.add(parseInt(m[1], 10)); });
  let pref = '';
  for (let r = 0; r < filas.length; r++) {
    const f = filas[r] || [];
    const celdas = filaTxt(f);
    if (!celdas.length) continue;
    const msp = celdas.join(' ').match(RE_SUBP);
    if (msp) { if (nSubp.size >= 2) pref = String(parseInt(msp[1], 10)).padStart(2, '0') + '.'; continue; }
    const t0 = normal(celdas[0]);
    // encabezado de partida
    let cod = null, des = '';
    if (/^(SUB\s*)?PARTIDA\b/.test(t0) && !/^(SUB\s*)?PARTIDAS\b/.test(t0)) {
      esSub = /^SUB/.test(t0);
      const resto = T(celdas[0]).replace(/^(sub\s*)?partida\s*:?\s*/i, '');
      const m = resto.match(/^([\w.]+)\s*(.*)$/);
      if (m && m[1]) { cod = m[1]; des = m[2]; }
      if (!cod || (!RE_ITEM.test(cod) && !esSub)) { cod = T(celdas[1]); des = T(celdas[2] || ''); }
      else if (!des) des = T(celdas[1] || '');
      cod = normItem(cod);
    } else if (celdas.length <= 3 && /^\d{1,4}(\.\d{1,4})+$/.test(normItem(celdas[0])) && !f.some((x) => typeof x === 'number' && x !== 0 && T(x) !== celdas[0])) {
      cod = normItem(celdas[0]); des = T(celdas[1] || ''); esSub = false;
    } else {
      const m = T(celdas[0]).match(/^(\d{1,4}(?:\.\d{1,4})+)\s+([A-ZÁÉÍÓÚÑ].{3,})$/i);
      if (m && celdas.length <= 2) { cod = m[1]; des = m[2]; esSub = false; }
    }
    if (cod && !esSub && pref && RE_ITEM.test(cod)) cod = pref + cod;
    if (cod) {
      const mr = celdas.join(' ').match(/REND[^\d]*([\d.,]+)/i);
      if (mr && !esSub) rend[cod] = num(mr[1]);
      if (esSub) { cur = { sub: true, cod, des, filas: [] }; subs[normal(cod)] = cur; subs[normal(des)] = cur; }
      else { cur = { sub: false, cod, des, filas: [] }; acu[cod] = cur.filas; }
      tipo = null;
      continue;
    }
    if (!cur) continue;
    // rendimiento
    if (celdas.some((c) => normal(c).startsWith('REND'))) {
      const txt = celdas.join(' ');
      const m = txt.match(/M\.?\s*O\.?\s*:?\s*([\d.,]+)/i) || txt.match(/REND[^\d]*([\d.,]+)/i);
      if (m && !cur.sub) rend[cur.cod] = num(m[1]);
      continue;
    }
    // títulos de columnas
    const hs = f.map(H);
    if (hs.some((h) => h.includes('CANTIDAD')) && hs.some((h) => h.startsWith('PRECIO') || h === 'PU') ) {
      cols = {};
      hs.forEach((h, c) => {
        if (!h) return;
        if (cols.und === undefined && ['UNIDAD', 'UND', 'UNID', 'UM'].includes(h)) cols.und = c;
        else if (cols.cuad === undefined && h.startsWith('CUAD')) cols.cuad = c;
        else if (cols.cant === undefined && h.startsWith('CANT')) cols.cant = c;
        else if (cols.precio === undefined && (h.startsWith('PRECIO') || h === 'PU')) cols.precio = c;
        else if (cols.parcial === undefined && h.startsWith('PARCIAL')) cols.parcial = c;
        else if (cols.desc === undefined && (h.includes('DESCRIP') || h.includes('RECURSO') || h.includes('INSUMO'))) cols.desc = c;
        else if (cols.cod === undefined && h.startsWith('COD')) cols.cod = c;
      });
      continue;
    }
    // sección
    if (celdas.length <= 2) {
      const e = ETIQ(celdas[0]);
      if (e) { tipo = e; continue; }
    }
    // fila de recurso
    let nombre = '', und = '', cuad = null, cant = null, precio = 0;
    if (cols && cols.cant !== undefined) {
      nombre = T(f[cols.desc ?? 1]);
      und = T(f[cols.und]);
      if (cols.cod !== undefined && cols.desc === undefined) nombre = T(f[cols.cod + 1]);
      const vCant = f[cols.cant];
      if (!esNumero(vCant)) continue;
      cant = num(vCant);
      cuad = cols.cuad !== undefined && esNumero(f[cols.cuad]) ? num(f[cols.cuad]) : null;
      precio = cols.precio !== undefined ? num(f[cols.precio]) : 0;
    } else {
      const nums = f.map((x, c) => [x, c]).filter(([x]) => typeof x === 'number');
      if (nums.length < 2) continue;
      const textos = f.map((x, c) => [x, c]).filter(([x]) => typeof x === 'string' && x.trim());
      if (!textos.length) continue;
      nombre = T(textos.find(([x]) => x.trim().length > 2)?.[0] || textos[0][0]);
      und = T(textos.length > 1 ? textos[textos.length - 1][0] : '');
      const v = nums.map(([x]) => x);
      if (v.length >= 4) { cuad = v[v.length - 4]; cant = v[v.length - 3]; precio = v[v.length - 2]; }
      else if (v.length === 3) { cant = v[0]; precio = v[1]; }
      else { cant = v[0]; precio = v[1]; }
    }
    nombre = quitarCodigo(nombre);
    if (!nombre || ETIQ(nombre)) continue;
    if (/^(COSTO UNITARIO|TOTAL|SUBTOTAL)/i.test(nombre)) continue;
    const tp = tipoPorUnidad(und, tipo || null);
    cur.filas.push({ tipo: tp, nombre, und, cuadrilla: cuad, cantidad: cant || 0, precio });
  }
  // reemplazar subpartidas por sus insumos (ModImportar.ExpandirSubpartidas)
  let nExp = 0, nNo = 0;
  for (const [it, filasA] of Object.entries(acu)) {
    const nuevas = [];
    for (const f of filasA) {
      if (f.tipo !== 'SP') { nuevas.push(f); continue; }
      const s = subs[normal(f.nombre)] || subs[normal(f.nombre.split(' ')[0])];
      if (s && s.filas.length) {
        nExp++;
        for (const g of s.filas) if (g.tipo !== '%MO') nuevas.push({ ...g, cantidad: g.cantidad * f.cantidad, cuadrilla: null });
      } else { nNo++; nuevas.push(f); }
    }
    acu[it] = nuevas;
  }
  if (nExp) avisos.push(`${nExp} subpartida(s) reemplazada(s) por sus insumos.`);
  if (nNo) avisos.push(`${nNo} subpartida(s) sin análisis: quedan como tipo SP (no se controlan).`);
  return { acu, rendimientos: rend, avisos };
}

/** Aplica un A.C.U. analizado a la obra. Devuelve un resumen. */
export function aplicarACU(obra, res, { reemplazar = true, actualizarPrecios = true } = {}) {
  const porClave = new Map(obra.partidas.filter((p) => p.tipo === 'P').map((p) => [claveItem(p.item), p.item]));
  const porItem = new Map(obra.partidas.map((p) => [p.item, p]));
  const insumos = new Map(obra.insumos.map((i) => [normal(i.descripcion), i]));
  let nP = 0, nF = 0, nNuevos = 0, nPrecios = 0;
  const sinPartida = [];
  const pref = { MO: 'MO', MAT: 'MT', EQ: 'EQ', SC: 'SC', SP: 'SP' };
  const sig = (tp) => {
    const p = pref[tp] || 'OT';
    let n = 0;
    for (const i of insumos.values()) if (i.codigo?.startsWith(p + '-')) n = Math.max(n, parseInt(i.codigo.slice(p.length + 1), 10) || 0);
    return `${p}-${String(n + 1).padStart(4, '0')}`;
  };
  const noAlm = String(obra.parametros?.no_almacen || '').split(';').map(normal).filter(Boolean);
  for (const [cod, filas] of Object.entries(res.acu)) {
    const item = porItem.has(cod) && porItem.get(cod).tipo === 'P' ? cod : porClave.get(claveItem(cod));
    if (!item) { sinPartida.push(cod); continue; }
    nP++;
    const nuevas = filas.map((f) => {
      nF++;
      if (f.tipo === '%MO') return { tipo: '%MO', descripcion: f.nombre, cantidad: f.cantidad };
      const k = normal(f.nombre);
      let i = insumos.get(k);
      if (!i) {
        i = { codigo: sig(f.tipo), descripcion: f.nombre.toUpperCase() === f.nombre ? f.nombre : f.nombre, und: f.und, tipo: f.tipo, precio: f.precio || 0, controlar: !noAlm.some((w) => k.includes(w)) };
        insumos.set(k, i);
        obra.insumos.push(i);
        nNuevos++;
      } else if (actualizarPrecios && f.precio && Math.abs((Number(i.precio) || 0) - f.precio) > 1e-9) { i.precio = f.precio; nPrecios++; }
      const fila = { tipo: f.tipo, insumo: i.descripcion, cantidad: f.cantidad };
      if (f.cuadrilla) fila.cuadrilla = f.cuadrilla;
      return fila;
    });
    obra.acu[item] = reemplazar ? nuevas : [...(obra.acu[item] || []), ...nuevas];
    const rd = res.rendimientos[cod];
    if (rd) porItem.get(item).rendimiento = rd;
  }
  return { partidas: nP, filas: nF, nuevos: nNuevos, precios: nPrecios, sinPartida };
}

// ------------------------------------------------------------------ tablas planas
export function leerTabla(filas, defs, minimo = 2) {
  const enc = buscarEncabezado(filas, defs, minimo);
  if (!enc) return null;
  const out = [];
  for (let r = enc.fila + 1; r < filas.length; r++) {
    const f = filas[r] || [];
    if (!f.some((x) => T(x))) continue;
    const o = { _fila: r + 1 };
    for (const [k, c] of Object.entries(enc.map)) o[k] = f[c];
    out.push(o);
  }
  return { filas: out, map: enc.map, filaEnc: enc.fila, encabezado: filas[enc.fila] };
}
const eq = (...xs) => (t) => xs.includes(t);
export const DEF_METRADOS = { fecha: eq('FECHA', 'DIA', 'FEC'), item: eq('ITEM', 'PARTIDA', 'CODIGO', 'COD'), metrado: (t) => t.startsWith('METRADO') || t.startsWith('CANT') };
export const DEF_MOVS = {
  fecha: eq('FECHA'), tipo: eq('TIPO', 'MOVIMIENTO'), codigo: eq('CODIGO', 'COD', 'CODIGOINSUMO'), descripcion: (t) => t.startsWith('DESCRIP') || t === 'INSUMO', und: eq('UND', 'UNIDAD', 'UM'),
  cantidad: (t) => t.startsWith('CANT'), pu: eq('PU', 'PRECIO', 'PRECIOUNITARIO', 'PUNIT'), documento: eq('DOCUMENTO', 'TIPODOCUMENTO', 'DOC'), numero: (t) => t.startsWith('NUMERO') || t === 'NDOCUMENTO' || t === 'N',
  procedencia: (t) => t.startsWith('PROCEDENCIA') || t.startsWith('PROVEEDOR'), partida: eq('PARTIDA', 'ITEM', 'DESTINO'), pedido_por: (t) => t.startsWith('PEDIDO'), vale: (t) => t.startsWith('VALE') || t === 'PECOSA', estado: eq('ESTADO'), oc_os: eq('OCOS', 'OC', 'OS'),
};
export const DEF_INSUMOS = { descripcion: (t) => t.startsWith('DESCRIP') || t === 'INSUMO' || t === 'RECURSO', und: eq('UND', 'UNIDAD', 'UM'), tipo: eq('TIPO'), precio: (t) => t.startsWith('PRECIO') || t === 'PU' || t === 'COSTO' };

export function importarMetrados(obra, filas, { sumar = false } = {}) {
  const t = leerTabla(filas, DEF_METRADOS, 3);
  if (!t) throw new Error('Se esperan columnas FECHA, ITEM y METRADO.');
  const porClave = new Map(obra.partidas.filter((p) => p.tipo === 'P').map((p) => [claveItem(p.item), p.item]));
  let n = 0;
  const errores = [];
  for (const f of t.filas) {
    const fe = aISO(f.fecha);
    const it = porClave.get(claveItem(textoItem(f.item) || T(f.item)));
    const m = num(f.metrado);
    if (!fe) { errores.push(`Fila ${f._fila}: fecha no válida`); continue; }
    if (!it) { errores.push(`Fila ${f._fila}: la partida ${T(f.item)} no existe`); continue; }
    obra.metrados[fe] = obra.metrados[fe] || {};
    obra.metrados[fe][it] = r2((sumar ? Number(obra.metrados[fe][it]) || 0 : 0) + m);
    n++;
  }
  return { n, errores };
}
export function importarMovimientos(obra, filas) {
  const t = leerTabla(filas, DEF_MOVS, 3);
  if (!t || t.map.fecha === undefined || t.map.cantidad === undefined) throw new Error('Se esperan al menos las columnas FECHA, TIPO, CODIGO (o DESCRIPCION) y CANTIDAD.');
  const porCod = new Map(obra.almacen.map((a) => [normal(a.codigo), a]));
  const porDesc = new Map(obra.almacen.map((a) => [normal(a.descripcion), a]));
  let n = 0, nuevos = 0;
  const errores = [];
  for (const f of t.filas) {
    const fe = aISO(f.fecha);
    if (!fe) { errores.push(`Fila ${f._fila}: fecha no válida`); continue; }
    let tipo = normal(T(f.tipo));
    tipo = tipo.startsWith('SAL') || tipo === 'S' || tipo.startsWith('EGRE') ? 'SALIDA' : tipo.startsWith('DEV') ? 'DEVOLUCION' : 'INGRESO';
    let a = porCod.get(normal(T(f.codigo))) || porDesc.get(normal(T(f.descripcion)));
    if (!a) {
      if (!T(f.descripcion) && !T(f.codigo)) { errores.push(`Fila ${f._fila}: falta el código o la descripción`); continue; }
      a = { codigo: T(f.codigo) || siguienteCodigoAlmacen(obra), descripcion: T(f.descripcion) || T(f.codigo), und: T(f.und) || 'UND', grupo: 'MATERIALES', pu: num(f.pu), insumo_acu: '', factor: 1, observacion: 'Creado al importar movimientos' };
      const ins = obra.insumos.find((i) => normal(i.descripcion) === normal(a.descripcion));
      if (ins) a.insumo_acu = ins.descripcion;
      obra.almacen.push(a);
      porCod.set(normal(a.codigo), a); porDesc.set(normal(a.descripcion), a);
      nuevos++;
    }
    obra.movimientos.push({ id: uid('M'), fecha: fe, tipo, documento: T(f.documento), numero: T(f.numero), oc_os: T(f.oc_os), procedencia: T(f.procedencia), partida: T(f.partida), codigo: a.codigo, cantidad: num(f.cantidad), pedido_por: T(f.pedido_por), vale: T(f.vale), observacion: '', estado: normal(T(f.estado)) === 'POR CONFIRMAR' ? 'POR CONFIRMAR' : 'CONFIRMADO', origen: 'IMPORTADO' });
    n++;
  }
  return { n, nuevos, errores };
}
export function siguienteCodigoAlmacen(obra, pref = 'M') {
  let n = 0;
  for (const a of obra.almacen) { const m = String(a.codigo).match(new RegExp('^' + pref + '-(\\d+)$')); if (m) n = Math.max(n, Number(m[1])); }
  return `${pref}-${String(n + 1).padStart(2, '0')}`;
}
export function importarPrecios(obra, filas) {
  const t = leerTabla(filas, DEF_INSUMOS, 2);
  if (!t || t.map.descripcion === undefined || t.map.precio === undefined) throw new Error('Se esperan columnas DESCRIPCION y PRECIO.');
  const porDesc = new Map(obra.insumos.map((i) => [normal(i.descripcion), i]));
  let act = 0, nuevos = 0;
  for (const f of t.filas) {
    const d = T(f.descripcion);
    if (!d) continue;
    const i = porDesc.get(normal(d));
    if (i) { if (Number(i.precio) !== num(f.precio)) { i.precio = num(f.precio); act++; } }
    else {
      const tipo = ['MO', 'MAT', 'EQ', 'SC'].includes(normal(T(f.tipo))) ? normal(T(f.tipo)) : tipoPorUnidad(T(f.und), null);
      obra.insumos.push({ codigo: 'NV-' + String(obra.insumos.length + 1).padStart(4, '0'), descripcion: d, und: T(f.und), tipo, precio: num(f.precio), controlar: true });
      nuevos++;
    }
  }
  return { actualizados: act, nuevos };
}
function mesDeEncabezado(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 7);
  if (typeof v === 'number' && v > 20000) return mesDe(aISO(v));
  const s = normal(T(v));
  let m = s.match(/^(\d{4})[-/](\d{1,2})$/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})[-/](\d{4})$/);
  if (m) return `${m[2]}-${m[1].padStart(2, '0')}`;
  const i = MESES.findIndex((x) => s.startsWith(x) || s.startsWith(x.slice(0, 3)) || (x === 'SETIEMBRE' && s.startsWith('SEP')));
  const y = s.match(/(\d{4})|(\d{2})$/);
  if (i >= 0 && y) return `${y[1] || '20' + y[2]}-${String(i + 1).padStart(2, '0')}`;
  return null;
}
export function importarCronograma(obra, filas) {
  for (let r = 0; r < Math.min(filas.length, 40); r++) {
    const f = filas[r] || [];
    const ci = f.findIndex((x) => ['ITEM', 'PARTIDA', 'CODIGO'].includes(H(x)));
    if (ci < 0) continue;
    const meses = [];
    f.forEach((x, c) => { const ym = c > ci ? mesDeEncabezado(x) : null; if (ym) meses.push([c, ym]); });
    if (!meses.length) continue;
    const porClave = new Map(obra.partidas.filter((p) => p.tipo === 'P').map((p) => [claveItem(p.item), p.item]));
    const lista = [...new Set([...(obra.cronograma.meses || []), ...meses.map((x) => x[1])])].sort();
    const montos = {};
    for (const [it, arr] of Object.entries(obra.cronograma.montos || {})) montos[it] = lista.map((ym) => arr[(obra.cronograma.meses || []).indexOf(ym)] || 0);
    let n = 0;
    for (let k = r + 1; k < filas.length; k++) {
      const g = filas[k] || [];
      const it = porClave.get(claveItem(textoItem(g[ci]) || T(g[ci])));
      if (!it) continue;
      montos[it] = montos[it] || lista.map(() => 0);
      for (const [c, ym] of meses) montos[it][lista.indexOf(ym)] = r2(num(g[c]));
      n++;
    }
    obra.cronograma = { meses: lista, montos };
    return { n, meses: meses.length };
  }
  throw new Error('No se encontró una fila de títulos con ITEM y columnas de meses (p. ej. 2026-10 u OCTUBRE 2026).');
}

// ------------------------------------------------------------------ JSON de la valorización (herramientas/exportar_valorizacion_json.py)
export function obraDesdeValorizacionJSON(val) {
  if (!val?.valorizacion?.partidas) throw new Error('El JSON no tiene el formato de la valorización exportada (falta valorizacion.partidas).');
  const o = val.obra || {};
  const u = o.ubicacion || {};
  const pl = val.plazo || {};
  const ind = (val.presupuesto || []).filter((p) => typeof p.pct_cd === 'number' && !/^(COSTO DIRECTO|TOTAL COSTO INDIRECTO|PRESUPUESTO TOTAL)/.test(p.concepto) && !p.concepto.startsWith('PIM')).map((p) => ({ concepto: p.concepto, monto: p.monto }));
  const partidas = [], anterior = {}, actual = {};
  for (const p of val.valorizacion.partidas) {
    const esP = p.tipo === 'partida';
    partidas.push({ item: p.item, tipo: esP ? 'P' : 'T', descripcion: T(p.descripcion), und: esP ? T(p.und) : '', metrado: esP ? p.metrado : null, pu: esP ? p.pu : null, rendimiento: null, origen: 'ET' });
    if (esP) {
      if (p.anterior?.metrado) anterior[p.item] = p.anterior.metrado;
      if (p.actual?.metrado) actual[p.item] = p.actual.metrado;
    }
  }
  const inf = val.informe || {};
  const cr = val.cronograma_valorizado;
  const items = new Set(partidas.filter((p) => p.tipo === 'P').map((p) => p.item));
  const cron = { meses: [], montos: {} };
  if (cr?.meses) {
    cron.meses = cr.meses.map((m) => mesDeEncabezado(m)).filter(Boolean);
    for (const p of cr.partidas || []) if (items.has(p.item)) cron.montos[p.item] = (p.por_mes || []).map((x) => r2(x || 0));
  }
  const total = (val.presupuesto || []).find((p) => p.concepto === 'PRESUPUESTO TOTAL')?.monto || 0;
  const dev = {};
  (val.curva_s || []).forEach((c, i) => { if (c.devengado_pct_acum !== null && c.devengado_pct_acum !== undefined && cron.meses[i]) dev[cron.meses[i]] = r2(c.devengado_pct_acum * total); });
  return {
    id: uid('obra-'),
    datos: { nombre: T(o.nombre).replace(/^[“"]|[”"]$/g, ''), nombre_corto: T(o.nombre_corto), cui: T(o.cui), meta: T(o.meta), entidad: T(o.unidad_ejecutora), dependencia: T(o.area_responsable), modalidad: T(o.modalidad), fuente_financiamiento: T(o.fuente_financiamiento), region: T(u.region), provincia: T(u.provincia), distrito: T(u.distrito), localidad: T(u.localidad), fecha_inicio: pl.inicio || null, plazo_dias: pl.dias || 0, ampliacion_dias: (pl.ampliaciones || []).reduce((s, a) => s + (a.dias || 0), 0) },
    responsables: (val.responsables || []).map((r) => ({ cargo: r.cargo, nombre: r.nombre, profesion: r.profesion || '' })),
    documentos: val.documentos || [], feriados: val.feriados || [], indirectos: ind, partidas,
    anterior_inicial: anterior,
    periodos: [{ numero: Number(inf.valorizacion_n) || 1, del: inf.periodo_del || pl.inicio, al: inf.periodo_al || pl.inicio, estado: 'ABIERTO', manual: actual, cerrado: null }],
    cronograma: cron, financiero: { devengado_acumulado: dev },
  };
}
