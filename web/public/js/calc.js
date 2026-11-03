// Motor de cálculo: presupuesto, A.C.U., consumo de insumos, estados del control diario,
// valorización (4.3, 4.4, 4.5, Curva S), almacén (kardex, 4.10, 4.11, 4.12) y equipos (4.15).
// Reproduce la lógica de los libros de Excel/VBA (ModActualizar.Calcular, ModReporte, ModProcesar).
// Sin dependencias del navegador: se prueba con node --test.

import {
  r2, r4, redondear, normal, cmpItem, sumarDias, difDias, diaSemana, mesDe, mesesEntre,
  diasDelMes, finDeMes, redondearCantidad,
} from './util.js';

export const ESTADOS = {
  VALIDO: { cod: 'VALIDO', txt: 'VÁLIDO', clase: 'e-valido', orden: 0 },
  REVISAR: { cod: 'REVISAR', txt: 'REVISAR / POR CONFIRMAR', clase: 'e-revisar', orden: 1 },
  SUPERA_REND: { cod: 'SUPERA_REND', txt: 'SUPERA RENDIMIENTO', clase: 'e-rend', orden: 2 },
  SUPERA_MET: { cod: 'SUPERA_MET', txt: 'SUPERA METRADO DEL E.T.', clase: 'e-met', orden: 3 },
  NO_VALORIZAR: { cod: 'NO_VALORIZAR', txt: 'NO VALORIZAR', clase: 'e-novalorizar', orden: 4 },
  FUERA_PLAZO: { cod: 'FUERA_PLAZO', txt: 'FUERA DEL PLAZO', clase: 'e-fuera', orden: 5 },
};

export const TIPOS_INSUMO = { MO: 'Mano de obra', MAT: 'Materiales', EQ: 'Equipos', SC: 'Subcontratos / servicios', '%MO': 'Herramientas (% M.O.)', SP: 'Subpartida' };
export const CONSUMIBLE = (t) => t === 'MO' || t === 'MAT' || t === 'EQ' || t === 'SC';
export const CONTROLABLE = (t) => t === 'MAT' || t === 'EQ' || t === 'SC';
const SIGNO = { INGRESO: 1, DEVOLUCION: 1, SALIDA: -1 };

/** Fecha de término (inicio + plazo + ampliaciones − 1). */
export function fechaTermino(d) {
  if (!d?.fecha_inicio) return null;
  const dias = (Number(d.plazo_dias) || 0) + (Number(d.ampliacion_dias) || 0);
  return dias > 0 ? sumarDias(d.fecha_inicio, dias - 1) : d.fecha_inicio;
}

/** Padre de un ítem: el título existente más largo que es prefijo. */
function buscarPadre(item, existe) {
  const seg = item.split('.');
  for (let n = seg.length - 1; n >= 1; n--) {
    const c = seg.slice(0, n).join('.');
    if (existe.has(c)) return c;
  }
  return null;
}

/**
 * Crea el motor de cálculo de una obra. Todo se calcula bajo demanda y se guarda en caché;
 * después de modificar la obra hay que crear un motor nuevo.
 */
export function crearMotor(obra) {
  const M = { obra };
  const P = obra.partidas || [];
  const cache = new Map();
  const memo = (k, fn) => {
    if (!cache.has(k)) cache.set(k, fn());
    return cache.get(k);
  };

  // ------------------------------------------------------------ índice del presupuesto
  const porItem = new Map();
  P.forEach((p) => porItem.set(p.item, p));
  const titulos = new Set(P.filter((p) => p.tipo === 'T').map((p) => p.item));
  const padre = new Map();
  const hijos = new Map();
  for (const p of P) {
    const pa = buscarPadre(p.item, titulos);
    padre.set(p.item, pa);
    if (!hijos.has(pa)) hijos.set(pa, []);
    hijos.get(pa).push(p.item);
  }
  M.porItem = porItem;
  M.padre = padre;
  M.hijos = hijos;
  M.partidasP = P.filter((p) => p.tipo === 'P');
  M.raices = hijos.get(null) || [];
  M.nivel = (item) => {
    let n = 1, x = padre.get(item);
    while (x) { n++; x = padre.get(x); }
    return n;
  };
  /** Subpresupuesto (título raíz) de un ítem. */
  M.raiz = (item) => {
    let x = item;
    while (padre.get(x)) x = padre.get(x);
    return x;
  };

  M.parcial = (p) => (p.tipo === 'P' ? r2((Number(p.metrado) || 0) * (Number(p.pu) || 0)) : 0);

  /** Suma hacia arriba un valor por partida: devuelve Map item → total (títulos incluidos). */
  M.sumarArbol = (valorDe) => {
    const t = new Map();
    for (let i = P.length - 1; i >= 0; i--) {
      const p = P[i];
      const v = p.tipo === 'P' ? valorDe(p) || 0 : t.get(p.item) || 0;
      t.set(p.item, r2(v));
      const pa = padre.get(p.item);
      if (pa) t.set(pa, (t.get(pa) || 0) + r2(v));
    }
    return t;
  };

  M.parciales = () => memo('parciales', () => M.sumarArbol(M.parcial));
  M.costoDirecto = () => memo('cd', () => r2(M.partidasP.reduce((s, p) => s + M.parcial(p), 0)));
  M.indirectos = () =>
    memo('ind', () => {
      const cd = M.costoDirecto();
      return (obra.indirectos || []).map((x) => ({
        concepto: x.concepto,
        monto: r2(x.monto),
        ratio: cd > 0 ? (Number(x.monto) || 0) / cd : 0,
      }));
    });
  M.totalIndirecto = () => r2(M.indirectos().reduce((s, x) => s + x.monto, 0));
  M.presupuestoTotal = () => r2(M.costoDirecto() + M.totalIndirecto());
  M.subpresupuestos = () =>
    memo('subp', () => {
      const par = M.parciales();
      return M.raices.map((it) => ({ item: it, descripcion: porItem.get(it)?.descripcion || '', presupuesto: par.get(it) || 0 }));
    });

  // ------------------------------------------------------------ insumos y A.C.U.
  const ins = new Map();
  (obra.insumos || []).forEach((i) => ins.set(normal(i.descripcion), i));
  M.insumo = (nombre) => ins.get(normal(nombre));
  M.insumos = obra.insumos || [];

  /** Filas del A.C.U. con precio y parcial calculados. */
  M.acu = (item) =>
    memo('acu:' + item, () => {
      const filas = (obra.acu && obra.acu[item]) || [];
      const out = filas.map((f) => {
        const i = f.tipo === '%MO' ? null : M.insumo(f.insumo);
        const precio = f.tipo === '%MO' ? 0 : i ? Number(i.precio) || 0 : Number(f.precio) || 0;
        return {
          ...f,
          descripcion: f.tipo === '%MO' ? f.descripcion || 'HERRAMIENTAS MANUALES' : f.insumo,
          und: f.tipo === '%MO' ? '%mo' : i?.und || f.und || '',
          codigo: i?.codigo || '',
          precio,
          parcial: f.tipo === '%MO' ? 0 : (Number(f.cantidad) || 0) * precio,
          existe: f.tipo === '%MO' || !!i,
        };
      });
      const mo = out.filter((f) => f.tipo === 'MO').reduce((s, f) => s + f.parcial, 0);
      out.forEach((f) => {
        if (f.tipo === '%MO') {
          f.precio = mo;
          f.parcial = ((Number(f.cantidad) || 0) / 100) * mo;
        }
      });
      return out;
    });
  M.puACU = (item) => r2(M.acu(item).reduce((s, f) => s + f.parcial, 0));
  M.resumenACU = (item) => {
    const r = { MO: 0, MAT: 0, EQ: 0, SC: 0, '%MO': 0, SP: 0 };
    M.acu(item).forEach((f) => (r[f.tipo] = (r[f.tipo] || 0) + f.parcial));
    return r;
  };

  /** Consumo de insumos para una lista de líneas {item, metrado}. */
  M.consumo = (lineas, { incluirMO = true } = {}) => {
    const mapa = new Map();
    for (const l of lineas) {
      const m = Number(l.metrado) || 0;
      if (!m) continue;
      for (const f of M.acu(l.item)) {
        if (!CONSUMIBLE(f.tipo)) continue;
        if (!incluirMO && f.tipo === 'MO') continue;
        const k = normal(f.descripcion);
        if (!mapa.has(k)) mapa.set(k, { insumo: f.descripcion, codigo: f.codigo, und: f.und, tipo: f.tipo, precio: f.precio, cantidad: 0, detalle: [] });
        const x = mapa.get(k);
        const q = m * (Number(f.cantidad) || 0);
        x.cantidad += q;
        x.detalle.push({ item: l.item, metrado: m, unitario: Number(f.cantidad) || 0, cantidad: q });
      }
    }
    const orden = { MAT: 0, EQ: 1, SC: 2, MO: 3 };
    return [...mapa.values()]
      .map((x) => ({ ...x, monto: x.cantidad * x.precio }))
      .sort((a, b) => (orden[a.tipo] ?? 9) - (orden[b.tipo] ?? 9) || a.insumo.localeCompare(b.insumo));
  };
  /** Explosión de insumos de todo el presupuesto (metrado del E.T. × A.C.U.). */
  M.explosion = () => memo('explosion', () => M.consumo(M.partidasP.map((p) => ({ item: p.item, metrado: p.metrado }))));

  // ------------------------------------------------------------ metrados diarios
  const diario = memo('diario', () => {
    const d = new Map();
    const fechas = Object.keys(obra.metrados || {}).sort();
    for (const f of fechas) {
      for (const [it, v] of Object.entries(obra.metrados[f])) {
        const x = Number(v) || 0;
        if (!x) continue;
        if (!d.has(it)) d.set(it, []);
        d.get(it).push([f, x]);
      }
    }
    return d;
  });
  M.fechasConMetrado = () => Object.keys(obra.metrados || {}).filter((f) => Object.values(obra.metrados[f]).some((v) => Number(v) > 0)).sort();
  M.metradoDia = (f, item) => Number(obra.metrados?.[f]?.[item]) || 0;
  M.sumaDiaria = (item, desde, hasta) => {
    let s = 0;
    for (const [f, v] of diario.get(item) || []) if ((!desde || f >= desde) && (!hasta || f <= hasta)) s += v;
    return redondear(s, 6);
  };
  M.diasDe = (item) => diario.get(item) || [];

  // ------------------------------------------------------------ disponibilidad de insumos (ingresos / O/S)
  const almacenPorCod = new Map((obra.almacen || []).map((a) => [a.codigo, a]));
  M.itemAlmacen = (cod) => almacenPorCod.get(cod);
  M.disponibilidad = () =>
    memo('dispo', () => {
      const ins = new Map(); // normal(insumo) → {any, conf, vig, abierto, cantidad}
      const par = new Map(); // item → {any, conf}
      const reg = (k, f, est, fin) => {
        if (!ins.has(k)) ins.set(k, { any: null, conf: null, finMax: null, abierto: false, cantidad: 0 });
        const x = ins.get(k);
        if (!x.any || f < x.any) x.any = f;
        if (est !== 'POR CONFIRMAR' && (!x.conf || f < x.conf)) x.conf = f;
        if (fin) { if (!x.finMax || fin > x.finMax) x.finMax = fin; } else x.abierto = true;
        return x;
      };
      for (const m of obra.movimientos || []) {
        if (m.tipo !== 'INGRESO' || !m.fecha) continue;
        const a = almacenPorCod.get(m.codigo);
        if (a?.insumo_acu) {
          const x = reg(normal(a.insumo_acu), m.fecha, m.estado, null);
          x.cantidad += (Number(m.cantidad) || 0) * (Number(a.factor) || 1);
        }
        if (m.partida_cubre) {
          const k = m.partida_cubre;
          if (!par.has(k)) par.set(k, { any: null, conf: null });
          const x = par.get(k);
          if (!x.any || m.fecha < x.any) x.any = m.fecha;
          if (m.estado !== 'POR CONFIRMAR' && (!x.conf || m.fecha < x.conf)) x.conf = m.fecha;
        }
      }
      for (const e of obra.equipos || []) {
        const f = e.inicio || e.os_fecha;
        if (e.insumo_acu && f) {
          const x = reg(normal(e.insumo_acu), f, e.estado, e.fin_os || null);
          x.cantidad += Number(e.cantidad) || 0;
        }
      }
      for (const x of ins.values()) x.vig = !x.abierto && x.finMax ? x.finMax : null;
      return { ins, par };
    });

  const palabrasNoAlmacen = () =>
    String(obra.parametros?.no_almacen || '')
      .split(';')
      .map((s) => normal(s))
      .filter(Boolean);
  M.controlaInsumo = (f) => {
    if (!CONTROLABLE(f.tipo)) return false;
    const i = M.insumo(f.descripcion);
    if (i && i.controlar === false) return false;
    const n = normal(f.descripcion);
    return !palabrasNoAlmacen().some((w) => n.includes(w));
  };

  /** Participación (% del P.U.) de cada insumo controlado del A.C.U. */
  M.participacion = (item) =>
    memo('part:' + item, () => {
      const p = porItem.get(item);
      const pu = Number(p?.pu) || 0;
      return M.acu(item)
        .filter((f) => M.controlaInsumo(f))
        .map((f) => ({ insumo: f.descripcion, tipo: f.tipo, peso: pu > 0 ? r4((f.cantidad * f.precio) / pu) : 0 }));
    });

  const inicio = obra.datos?.fecha_inicio || null;
  const termino = fechaTermino(obra.datos || {});
  M.inicio = inicio;
  M.termino = termino;
  M.umbral = () => Number(obra.parametros?.umbral ?? 0.2);

  /** Acumulado de una partida hasta una fecha (incluida): anterior inicial + diarios. */
  M.acumuladoHasta = (item, f) => (Number(obra.anterior_inicial?.[item]) || 0) + M.sumaDiaria(item, null, f);

  /**
   * Estado de una partida en un día con metrado (como ModActualizar.Calcular):
   * faltan = Σ peso de los insumos sin ingreso ni O/S vigente a esa fecha;
   * porConfirmar = Σ peso de los que solo tienen ingresos «POR CONFIRMAR».
   */
  M.estadoDia = (item, f) => {
    const p = porItem.get(item);
    const m = M.metradoDia(f, item);
    if (!p || p.tipo !== 'P' || m <= 0) return null;
    const { ins, par } = M.disponibilidad();
    const pp = par.get(item);
    let faltan = 0, porConf = 0;
    const lFalta = [], lConf = [];
    for (const x of M.participacion(item)) {
      if (x.peso <= 0) continue;
      const d = ins.get(normal(x.insumo));
      let any = d?.any || null, conf = d?.conf || null;
      if (pp?.any && (!any || pp.any < any)) any = pp.any;
      if (pp?.conf && (!conf || pp.conf < conf)) conf = pp.conf;
      const vig = pp?.any ? null : d?.vig || null;
      if (!any || any > f || (vig && vig < f)) {
        faltan += x.peso;
        lFalta.push(x.insumo);
      } else if (!conf || conf > f) {
        porConf += x.peso;
        lConf.push(x.insumo);
      }
    }
    const motivos = [];
    let est = ESTADOS.VALIDO;
    const sube = (e) => { if (e.orden > est.orden) est = e; };
    if ((inicio && f < inicio) || (termino && f > termino)) { sube(ESTADOS.FUERA_PLAZO); motivos.push('Fecha fuera del plazo de ejecución'); }
    const rend = (Number(p.rendimiento) || 0) * (Number(p.cuadrillas) || 1);
    if (rend > 0 && m > rend * 1.0001) { sube(ESTADOS.SUPERA_REND); motivos.push(`Supera el rendimiento (máx. ${redondear(rend, 2)} ${p.und}/día)`); }
    const acu = M.acumuladoHasta(item, f);
    if (acu > (Number(p.metrado) || 0) * 1.0001 + 0.001) { sube(ESTADOS.SUPERA_MET); motivos.push(`Acumulado ${redondear(acu, 2)} supera el metrado del E.T. (${p.metrado})`); }
    if (faltan > 0) {
      if (faltan >= M.umbral()) { sube(ESTADOS.NO_VALORIZAR); motivos.push(`Sin O/S ni ingreso (${redondear(faltan * 100, 1)} % del P.U.): ${lFalta.join(', ')}`); }
      else { sube(ESTADOS.REVISAR); motivos.push(`Falta insumo menor (${redondear(faltan * 100, 1)} % del P.U.): ${lFalta.join(', ')}`); }
    }
    if (porConf > 0) { sube(ESTADOS.REVISAR); motivos.push('Por confirmar: ' + lConf.join(', ')); }
    return { estado: est, faltan, porConfirmar: porConf, motivos, metrado: m, acumulado: acu, faltantes: lFalta, porConfirmarLista: lConf };
  };

  /** Resumen de estados del mes por partida: el peor estado de sus días. */
  M.estadosMes = (ym) =>
    memo('estMes:' + ym, () => {
      const out = new Map();
      const desde = ym + '-01', hasta = finDeMes(ym);
      for (const [item, dias] of diario) {
        for (const [f] of dias) {
          if (f < desde || f > hasta) continue;
          const e = M.estadoDia(item, f);
          if (!e) continue;
          const prev = out.get(item);
          if (!prev || e.estado.orden > prev.estado.orden) out.set(item, { ...e, fecha: f });
        }
      }
      return out;
    });

  // ------------------------------------------------------------ valorización
  M.periodos = () => [...(obra.periodos || [])].sort((a, b) => (a.del < b.del ? -1 : 1));
  M.indicePeriodo = (numero) => M.periodos().findIndex((p) => String(p.numero) === String(numero));

  /** Metrado del periodo de una partida y de dónde sale. */
  M.metradoPeriodo = (per, item) => {
    if (per.cerrado) return { m: Number(per.cerrado[item]) || 0, fuente: 'cerrado' };
    if (per.manual && per.manual[item] !== undefined && per.manual[item] !== null && per.manual[item] !== '') return { m: Number(per.manual[item]) || 0, fuente: 'manual' };
    return { m: M.sumaDiaria(item, per.del, per.al), fuente: 'diario' };
  };

  /** Metrado y monto de cada partida por periodo (todos los periodos). */
  M.tablaPeriodos = () =>
    memo('tablaPer', () => {
      const pers = M.periodos();
      return pers.map((per) => {
        const met = new Map(), mon = new Map(), fte = new Map();
        let cd = 0;
        for (const p of M.partidasP) {
          const { m, fuente } = M.metradoPeriodo(per, p.item);
          met.set(p.item, m);
          fte.set(p.item, fuente);
          const v = r2(m * (Number(p.pu) || 0));
          mon.set(p.item, v);
          cd += v;
        }
        return { per, met, mon, fte, cd: r2(cd) };
      });
    });

  /** Valorización completa de un periodo (índice en M.periodos()). */
  M.valorizacion = (k) =>
    memo('val:' + k, () => {
      const tabla = M.tablaPeriodos();
      const t = tabla[k];
      if (!t) return null;
      const antMet = new Map(), antMon = new Map();
      for (const p of M.partidasP) {
        const a0 = Number(obra.anterior_inicial?.[p.item]) || 0;
        let am = a0, av = r2(a0 * (Number(p.pu) || 0));
        for (let j = 0; j < k; j++) { am += tabla[j].met.get(p.item) || 0; av += tabla[j].mon.get(p.item) || 0; }
        antMet.set(p.item, redondear(am, 6));
        antMon.set(p.item, r2(av));
      }
      const par = M.parciales();
      const sAnt = M.sumarArbol((p) => antMon.get(p.item));
      const sAct = M.sumarArbol((p) => t.mon.get(p.item));
      const filas = P.map((p) => {
        const parcial = par.get(p.item) || 0;
        const ant = sAnt.get(p.item) || 0, act = sAct.get(p.item) || 0, acu = r2(ant + act);
        const f = {
          item: p.item, tipo: p.tipo, descripcion: p.descripcion, und: p.und || '', nivel: M.nivel(p.item),
          metrado: p.tipo === 'P' ? Number(p.metrado) || 0 : null, pu: p.tipo === 'P' ? Number(p.pu) || 0 : null, parcial,
          ant: { m: null, s: ant }, act: { m: null, s: act }, acu: { m: null, s: acu }, sal: { m: null, s: r2(parcial - acu) },
          pctAct: parcial ? act / parcial : 0, pctAcu: parcial ? acu / parcial : 0, fuente: null, origen: p.origen || 'ET',
        };
        if (p.tipo === 'P') {
          f.ant.m = antMet.get(p.item);
          f.act.m = t.met.get(p.item);
          f.acu.m = redondear(f.ant.m + f.act.m, 6);
          f.sal.m = redondear(f.metrado - f.acu.m, 6);
          f.fuente = t.fte.get(p.item);
          f.excede = f.acu.m > f.metrado * 1.0001 + 0.001;
        }
        return f;
      });
      const cdPres = M.costoDirecto();
      const cdAnt = r2(M.partidasP.reduce((s, p) => s + antMon.get(p.item), 0));
      const cdAct = t.cd;
      const res = (pres, ant, act) => ({ presupuesto: pres, ant, act, acu: r2(ant + act), sal: r2(pres - ant - act), pctAct: pres ? act / pres : 0, pctAcu: pres ? (ant + act) / pres : 0 });
      const subp = M.subpresupuestos().map((s) => ({ ...s, ...res(s.presupuesto, sAnt.get(s.item) || 0, sAct.get(s.item) || 0) }));
      const ind = M.indirectos().map((x) => {
        let ant = 0;
        for (let j = 0; j < k; j++) ant += r2(x.ratio * tabla[j].cd);
        const a0 = r2(x.ratio * M.partidasP.reduce((s, p) => s + r2((Number(obra.anterior_inicial?.[p.item]) || 0) * (Number(p.pu) || 0)), 0));
        return { concepto: x.concepto, ratio: x.ratio, ...res(x.monto, r2(ant + a0), r2(x.ratio * cdAct)) };
      });
      const cd = res(cdPres, cdAnt, cdAct);
      const ti = res(M.totalIndirecto(), r2(ind.reduce((s, x) => s + x.ant, 0)), r2(ind.reduce((s, x) => s + x.act, 0)));
      const total = res(M.presupuestoTotal(), r2(cd.ant + ti.ant), r2(cd.act + ti.act));
      // Formato 10: por origen de la partida
      const f10 = {};
      for (const f of filas) if (f.tipo === 'P') {
        const o = f.origen || 'ET';
        if (!f10[o]) f10[o] = { presupuesto: 0, ant: 0, act: 0 };
        f10[o].presupuesto += f.parcial; f10[o].ant += f.ant.s; f10[o].act += f.act.s;
      }
      return { periodo: t.per, k, filas, cd, subp, ind, totalIndirecto: ti, total, f10, nPartidasMes: [...t.met.values()].filter((x) => x > 0).length };
    });

  // ------------------------------------------------------------ cronograma, comparado y curva S
  M.cronoMeses = () => obra.cronograma?.meses || [];
  M.programadoItemMes = (item, ym) => {
    const i = M.cronoMeses().indexOf(ym);
    if (i < 0) return 0;
    return Number(obra.cronograma?.montos?.[item]?.[i]) || 0;
  };
  M.programadoMes = (ym) => r2(M.partidasP.reduce((s, p) => s + M.programadoItemMes(p.item, ym), 0));

  /** Cuadro 4.5: programado del mes vs ejecutado, por partida. */
  M.comparado = (k) =>
    memo('cmp:' + k, () => {
      const v = M.valorizacion(k);
      if (!v) return null;
      const ym = mesDe(v.periodo.al);
      const prog = M.sumarArbol((p) => r2(M.programadoItemMes(p.item, ym)));
      const filas = v.filas.map((f) => {
        const pr = prog.get(f.item) || 0;
        return { ...f, programado: pr, ejecutado: f.act.s, diferencia: r2(f.act.s - pr), cumplimiento: pr ? f.act.s / pr : null };
      });
      const programado = M.programadoMes(ym);
      return { ym, filas, programado, ejecutado: v.cd.act, diferencia: r2(v.cd.act - programado), cumplimiento: programado ? v.cd.act / programado : null };
    });

  M.curvaS = () =>
    memo('curva', () => {
      const tabla = M.tablaPeriodos();
      const meses = new Set(M.cronoMeses());
      tabla.forEach((t) => meses.add(mesDe(t.per.al)));
      const lista = [...meses].sort();
      const cd = M.costoDirecto(), tot = M.presupuestoTotal();
      let pa = 0, ea = 0;
      const dev = obra.financiero?.devengado_acumulado || {};
      const ultimoEjec = tabla.length ? mesDe(tabla[tabla.length - 1].per.al) : null;
      return lista.map((ym, i) => {
        const prog = M.programadoMes(ym);
        pa += prog;
        const ejecMes = tabla.filter((t) => mesDe(t.per.al) === ym).reduce((s, t) => s + t.cd, 0);
        ea += ejecMes;
        const hayEjec = ultimoEjec && ym <= ultimoEjec;
        return {
          n: i + 1, ym, programado: r2(prog), progAcum: r2(pa), progPct: cd ? pa / cd : 0,
          ejecutado: hayEjec ? r2(ejecMes) : null, ejecAcum: hayEjec ? r2(ea) : null, ejecPct: hayEjec && cd ? ea / cd : null,
          devengado: dev[ym] ?? null, devPct: dev[ym] !== undefined && dev[ym] !== null && tot ? dev[ym] / tot : null,
        };
      });
    });

  // ------------------------------------------------------------ almacén
  M.almacen = obra.almacen || [];
  /** Movimientos ordenados con saldo después de cada uno (ModProcesar). */
  M.movimientos = () =>
    memo('movs', () => {
      const lista = (obra.movimientos || []).map((m, i) => ({ ...m, _i: i }));
      lista.sort((a, b) => (a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : (SIGNO[b.tipo] || 0) - (SIGNO[a.tipo] || 0) || a._i - b._i));
      const saldo = new Map();
      for (const m of lista) {
        const s = r4((saldo.get(m.codigo) || 0) + (SIGNO[m.tipo] || 0) * (Number(m.cantidad) || 0));
        saldo.set(m.codigo, s);
        m.saldo = s;
        m.negativo = s < -0.00005;
      }
      return lista;
    });
  M.saldoActual = (cod, hasta = null) => {
    let s = 0;
    for (const m of M.movimientos()) if (m.codigo === cod && (!hasta || m.fecha <= hasta)) s += (SIGNO[m.tipo] || 0) * (Number(m.cantidad) || 0);
    return r4(s);
  };
  M.stock = () =>
    memo('stock', () => {
      const s = new Map();
      for (const m of M.movimientos()) {
        if (!s.has(m.codigo)) s.set(m.codigo, { ingresos: 0, salidas: 0, saldo: 0, ultimo: null });
        const x = s.get(m.codigo);
        if (SIGNO[m.tipo] > 0) x.ingresos += Number(m.cantidad) || 0;
        else x.salidas += Number(m.cantidad) || 0;
        x.saldo = m.saldo;
        x.ultimo = m.fecha;
      }
      return s;
    });
  M.kardex = (cod, desde = null, hasta = null) => {
    const a = almacenPorCod.get(cod);
    const pu = Number(a?.pu) || 0;
    let saldoIni = 0;
    const filas = [];
    for (const m of M.movimientos()) {
      if (m.codigo !== cod) continue;
      if (desde && m.fecha < desde) { saldoIni = m.saldo; continue; }
      if (hasta && m.fecha > hasta) continue;
      const ent = SIGNO[m.tipo] > 0 ? Number(m.cantidad) || 0 : 0;
      const sal = SIGNO[m.tipo] < 0 ? Number(m.cantidad) || 0 : 0;
      filas.push({ ...m, entrada: ent, salida: sal, pu, entradaS: r2(ent * pu), salidaS: r2(sal * pu), saldoS: r2(m.saldo * pu) });
    }
    return { item: a, pu, saldoInicial: saldoIni, filas };
  };
  /** 4.11 y 4.12: por ítem del almacén, saldo anterior, ingresos y salidas del mes, saldo. */
  M.resumenAlmacen = (ym) => {
    const ini = ym + '-01', fin = finDeMes(ym);
    return M.almacen.map((a) => {
      let ant = 0, ing = 0, sal = 0;
      for (const m of M.movimientos()) {
        if (m.codigo !== a.codigo) continue;
        const q = Number(m.cantidad) || 0, sg = SIGNO[m.tipo] || 0;
        if (m.fecha < ini) ant += sg * q;
        else if (m.fecha <= fin) { if (sg > 0) ing += q; else sal += q; }
      }
      const pu = Number(a.pu) || 0;
      const saldo = r4(ant + ing - sal);
      return { ...a, anterior: r4(ant), ingresos: r4(ing), salidas: r4(sal), saldo, pu, valorizado: r2(saldo * pu) };
    });
  };
  /** 4.10: movimiento diario del mes (solo días con movimiento). */
  M.movimientoDiario = (ym) => {
    const ini = ym + '-01', fin = finDeMes(ym);
    const dias = [...new Set(M.movimientos().filter((m) => m.fecha >= ini && m.fecha <= fin).map((m) => m.fecha))].sort();
    const filas = M.resumenAlmacen(ym).filter((r) => r.anterior || r.ingresos || r.salidas).map((r) => {
      const porDia = {};
      for (const m of M.movimientos()) {
        if (m.codigo !== r.codigo || m.fecha < ini || m.fecha > fin) continue;
        if (!porDia[m.fecha]) porDia[m.fecha] = { e: 0, s: 0 };
        if (SIGNO[m.tipo] > 0) porDia[m.fecha].e += Number(m.cantidad) || 0;
        else porDia[m.fecha].s += Number(m.cantidad) || 0;
      }
      return { ...r, porDia };
    });
    return { dias, filas };
  };

  // ------------------------------------------------------------ equipos
  M.horasParte = (p) => {
    const h = (s) => {
      const m = String(s || '').match(/^(\d{1,2})[:.h]?(\d{2})?/);
      return m ? Number(m[1]) + (Number(m[2]) || 0) / 60 : null;
    };
    let horario = 0, hay = false;
    for (const [a, b] of [[p.m_de, p.m_a], [p.t_de, p.t_a]]) {
      const x = h(a), y = h(b);
      if (x !== null && y !== null && y > x) { horario += y - x; hay = true; }
    }
    const horo = p.h_ini !== null && p.h_ini !== undefined && p.h_ini !== '' && p.h_fin !== null && p.h_fin !== undefined && p.h_fin !== '' ? Number(p.h_fin) - Number(p.h_ini) : null;
    const manual = p.horas !== null && p.horas !== undefined && p.horas !== '' ? Number(p.horas) : null;
    const valorizar = manual ?? (horo !== null && horo > 0 ? horo : hay ? horario : 0);
    const alertas = [];
    if (horo !== null && horo < 0) alertas.push('Horómetro final menor que el inicial');
    if (horo !== null && hay && Math.abs(horo - horario) > 0.5) alertas.push(`Horómetro (${r2(horo)} h) ≠ horario (${r2(horario)} h)`);
    return { horario: hay ? r2(horario) : null, horometro: horo !== null ? r2(horo) : null, valorizar: r2(valorizar), alertas };
  };
  /** 4.15: control de equipo mecánico del mes. */
  M.controlEquipos = (ym) => {
    const ini = ym + '-01', fin = finDeMes(ym);
    const nd = diasDelMes(ym);
    return (obra.equipos || []).map((e) => {
      const dias = Array(nd).fill(0), gal = Array(nd).fill(0);
      let ant = 0, mes = 0, galMes = 0;
      for (const p of obra.partes || []) {
        if (p.equipo !== e.codigo) continue;
        const h = M.horasParte(p).valorizar;
        if (p.fecha < ini) ant += h;
        else if (p.fecha <= fin) {
          const d = Number(p.fecha.slice(8, 10)) - 1;
          dias[d] += h; mes += h;
          gal[d] += Number(p.combustible) || 0; galMes += Number(p.combustible) || 0;
        }
      }
      const contratado = Number(e.cantidad) || 0;
      const acum = r2(ant + mes);
      return { ...e, dias: dias.map(r2), gal: gal.map(r2), anterior: r2(ant), mes: r2(mes), acumulado: acum, saldo: r2(contratado - acum), galMes: r2(galMes), valorizado: r2(mes * (Number(e.costo) || 0)), excede: contratado > 0 && acum > contratado + 0.001 };
    });
  };

  // ------------------------------------------------------------ generadores (A.C.U. × metrado)
  /** Propone salidas del almacén desde los metrados diarios (como ModImportar del almacén). */
  M.proponerSalidas = (desde, hasta, modo = 'arriba') => {
    const porInsumo = new Map();
    for (const a of M.almacen) {
      if (!a.insumo_acu) continue;
      const k = normal(a.insumo_acu);
      if (!porInsumo.has(k)) porInsumo.set(k, a);
    }
    const out = [];
    const sinEnlace = new Map();
    for (const f of Object.keys(obra.metrados || {}).sort()) {
      if (f < desde || f > hasta) continue;
      for (const [item, v] of Object.entries(obra.metrados[f])) {
        const m = Number(v) || 0;
        if (!m) continue;
        for (const r of M.acu(item)) {
          if (r.tipo !== 'MAT') continue;
          const a = porInsumo.get(normal(r.descripcion));
          const qAcu = m * (Number(r.cantidad) || 0);
          if (!a) {
            if (M.controlaInsumo(r)) sinEnlace.set(r.descripcion, (sinEnlace.get(r.descripcion) || 0) + qAcu);
            continue;
          }
          const q = qAcu / (Number(a.factor) || 1);
          out.push({ fecha: f, partida: item, codigo: a.codigo, descripcion: a.descripcion, und: a.und, insumo_acu: r.descripcion, teorico: q, cantidad: redondearCantidad(q, modo), metrado: m });
        }
      }
    }
    return { propuestas: out, sinEnlace: [...sinEnlace.entries()].map(([insumo, q]) => ({ insumo, cantidad: q })) };
  };
  /** Propone partes diarios de equipos desde los metrados (horas = metrado × cantidad del A.C.U.). */
  M.proponerPartes = (desde, hasta) => {
    const porInsumo = new Map();
    for (const e of obra.equipos || []) if (e.insumo_acu) porInsumo.set(normal(e.insumo_acu), e);
    const out = [];
    for (const f of Object.keys(obra.metrados || {}).sort()) {
      if (f < desde || f > hasta) continue;
      for (const [item, v] of Object.entries(obra.metrados[f])) {
        const m = Number(v) || 0;
        if (!m) continue;
        for (const r of M.acu(item)) {
          if (r.tipo !== 'EQ') continue;
          const e = porInsumo.get(normal(r.descripcion));
          if (!e) continue;
          out.push({ fecha: f, equipo: e.codigo, partida: item, trabajo: porItem.get(item)?.descripcion || '', metrado: m, horas: r2(m * (Number(r.cantidad) || 0)) });
        }
      }
    }
    return out;
  };

  // ------------------------------------------------------------ plazo
  M.avancePlazo = (f) => {
    if (!inicio || !termino) return null;
    const total = difDias(inicio, termino) + 1;
    const trans = Math.max(0, Math.min(total, difDias(inicio, f) + 1));
    return { total, transcurridos: trans, pct: total ? trans / total : 0 };
  };
  M.esLaborable = (f) => {
    const dl = obra.parametros?.dias_laborables || [1, 2, 3, 4, 5, 6];
    if (!dl.includes(diaSemana(f))) return false;
    return !(obra.feriados || []).some((x) => x.fecha === f);
  };

  return M;
}

/** Periodo siguiente: desde el día después del cierre hasta fin de ese mes (o el término). */
export function periodoSiguiente(obra, per) {
  const del = sumarDias(per.al, 1);
  let al = finDeMes(mesDe(del));
  const t = fechaTermino(obra.datos || {});
  if (t && al > t && del <= t) al = t;
  return { numero: (Number(per.numero) || 0) + 1, del, al, estado: 'ABIERTO', manual: {}, cerrado: null };
}

/** Periodo inicial de una obra: del inicio al fin de ese mes. */
export function periodoInicial(obra) {
  const ini = obra.datos?.fecha_inicio;
  if (!ini) return null;
  let al = finDeMes(mesDe(ini));
  const t = fechaTermino(obra.datos);
  if (t && al > t) al = t;
  return { numero: 1, del: ini, al, estado: 'ABIERTO', manual: {}, cerrado: null };
}

/** Cronograma lineal: reparte el parcial de cada partida en los meses del plazo según los días. */
export function cronogramaLineal(obra, motor) {
  const ini = obra.datos?.fecha_inicio, fin = fechaTermino(obra.datos || {});
  if (!ini || !fin) return null;
  const meses = mesesEntre(ini, fin);
  const total = difDias(ini, fin) + 1;
  const pesos = meses.map((ym) => {
    const a = ym + '-01' < ini ? ini : ym + '-01';
    const b = finDeMes(ym) > fin ? fin : finDeMes(ym);
    return (difDias(a, b) + 1) / total;
  });
  const montos = {};
  for (const p of motor.partidasP) {
    const par = motor.parcial(p);
    let acum = 0;
    montos[p.item] = pesos.map((w, i) => {
      if (i === pesos.length - 1) return r2(par - acum);
      const v = r2(par * w);
      acum += v;
      return v;
    });
  }
  return { meses, montos };
}

export { cmpItem };
