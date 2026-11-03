// Calculadora de insumos: metrado libre × A.C.U. → cantidades (con redondeo y edición) para pedidos y salidas.
import { S, cambio } from '../store.js';
import { esc, fmt, cant, num, normal, hoy, redondearCantidad, r2, uid } from '../util.js';
import { icono, ventana, aviso, htmlCampos, leerCampos } from '../ui.js';
import { rCalculadora } from '../reportes.js';
import { siguienteCodigoAlmacen } from '../importar.js';
import { cab, btnExport, enlazarExport, elegirPartidas, sinPresupuesto } from './comun.js';
import { refrescar, ir } from '../app.js';

function estado() {
  const o = S.obra;
  o.calculadora = o.calculadora || { lineas: [], ajustes: {}, modo: o.parametros.redondeo || 'arriba', paso: 1, mo: false, agrupar: 'insumo' };
  return o.calculadora;
}

/** Calcula las filas de insumos con redondeo y cantidades finales (ajustes del usuario). */
export function calcular() {
  const c = estado();
  const m = S.motor;
  const base = m.consumo(c.lineas.filter((l) => m.porItem.get(l.item)), { incluirMO: c.mo });
  const alm = new Map();
  for (const a of m.almacen) if (a.insumo_acu && !alm.has(normal(a.insumo_acu))) alm.set(normal(a.insumo_acu), a);
  return base.map((x) => {
    const k = normal(x.insumo);
    const red = redondearCantidad(x.cantidad, c.modo, Number(c.paso) || 1);
    const aj = c.ajustes[k];
    const editado = aj && Math.abs(aj.teorico - x.cantidad) < 1e-9;
    const final = editado ? aj.final : red;
    const a = alm.get(k);
    let txtAlm = '', qAlm = null;
    if (a) {
      qAlm = redondearCantidad(final / (Number(a.factor) || 1), c.modo === 'ninguno' || c.modo === 'dos' ? 'arriba' : c.modo);
      txtAlm = `${a.codigo}: ${cant(qAlm, 2)} ${a.und}` + ((Number(a.factor) || 1) !== 1 ? ` (1 ${a.und} = ${cant(a.factor, 4)} ${x.und})` : '');
    }
    return { ...x, k, redondeado: red, final, editado, alm: txtAlm, almItem: a || null, qAlm };
  });
}

export function render(cont, params) {
  const m = S.motor;
  if (!m.partidasP.length) { cont.innerHTML = cab('Calculadora de insumos', '') + sinPresupuesto(); return; }
  const c = estado();
  if (params.item && m.porItem.get(params.item) && !c.lineas.some((l) => l.item === params.item)) {
    c.lineas.push({ item: params.item, metrado: 1 });
    cambio();
  }
  const filas = calcular();
  const total = filas.reduce((s, x) => s + r2(x.final * x.precio), 0);
  const nEd = filas.filter((x) => x.editado).length;
  cont.innerHTML = `
    ${cab('Calculadora de insumos', 'Escriba cualquier metrado y obtenga los materiales, equipos y mano de obra según el A.C.U. Las cantidades se redondean (enteros, para el almacén) y puede <b>editar la cantidad final</b> a su conveniencia antes de exportar o registrar la salida.',
      `${btnExport('calc')}<button class="btn btn-primario" id="b-salida" ${filas.length ? '' : 'disabled'}>${icono('almacen')}Registrar salida de almacén</button>`)}
    <div class="rejilla-2" style="grid-template-columns: minmax(340px, 1fr) 2fr">
      <div class="panel" style="margin:0"><div class="panel-cab"><h2>Metrados</h2><div class="acciones"><button class="btn btn-sm btn-primario" id="b-add">${icono('mas')}Agregar partidas</button><button class="btn btn-sm" id="b-limpiar" ${c.lineas.length ? '' : 'disabled'}>Limpiar</button></div></div>
        <div class="panel-cuerpo">
          ${c.lineas.length ? `<table class="t"><thead><tr><th>Partida</th><th class="n">Metrado</th><th>Und</th><th></th></tr></thead><tbody id="t-lin">
          ${c.lineas.map((l, i) => { const p = m.porItem.get(l.item); return `<tr data-i="${i}"><td><div class="item mono" style="font-size:11px">${esc(l.item)}</div><div style="font-size:12px">${esc(p?.descripcion || 'PARTIDA NO EXISTE')}</div>${p && !m.acu(l.item).length ? '<span class="badge badge-alerta">sin A.C.U.</span>' : ''}</td><td><input class="n" data-k="metrado" value="${cant(l.metrado, 4)}" style="width:100px"></td><td>${esc(p?.und || '')}</td><td><button class="btn-icono peligro" data-del="${i}">${icono('borrar')}</button></td></tr>`; }).join('')}
          </tbody></table>` : `<div class="vacio" style="padding:20px"><p>Agregue una o varias partidas y escriba el metrado.</p></div>`}
          <div class="form-grid" style="grid-template-columns:1fr 1fr;margin-top:14px">
            <label class="campo"><span>Redondeo</span><select id="o-modo">${[['arriba', 'Entero hacia arriba'], ['normal', 'Entero más cercano'], ['abajo', 'Entero hacia abajo'], ['dos', '2 decimales'], ['ninguno', 'Sin redondeo (4 dec.)']].map(([k, t]) => `<option value="${k}" ${c.modo === k ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
            <label class="campo"><span>Múltiplo</span><select id="o-paso">${[1, 0.5, 0.25, 0.1, 5, 10, 50, 100].map((p) => `<option value="${p}" ${Number(c.paso) === p ? 'selected' : ''}>${p}</option>`).join('')}</select></label>
            <label class="chk" style="grid-column:1/-1"><input type="checkbox" id="o-mo" ${c.mo ? 'checked' : ''}> Incluir mano de obra (HH)</label>
          </div>
          ${c.lineas.length ? `<div class="acciones" style="margin-top:10px"><button class="btn btn-sm" id="b-met">${icono('metrados')}Registrar como metrado del día</button></div>` : ''}
        </div></div>
      <div class="panel" style="margin:0"><div class="panel-cab"><h2>Insumos calculados</h2><div class="acciones">${nEd ? `<span class="badge badge-alerta">${nEd} editado(s)</span><button class="btn btn-sm" id="b-reset">Restablecer</button>` : ''}<span class="badge badge-azul">Costo S/ ${fmt(total)}</span></div></div>
        <div class="tabla-cont libre" style="border:0">${filas.length ? `<table class="t" id="t-ins"><thead><tr><th>Insumo</th><th>Und</th><th>Tipo</th><th class="n">Según A.C.U.</th><th class="n">Redondeado</th><th class="n">Cantidad final</th><th class="n">Precio</th><th class="n">Costo S/</th><th>Almacén</th></tr></thead><tbody>
          ${filas.map((x) => `<tr data-k="${esc(x.k)}"><td>${esc(x.insumo)}</td><td>${esc(x.und)}</td><td><span class="badge">${x.tipo}</span></td><td class="n tenue" title="${esc(x.detalle.map((d) => `${d.item}: ${cant(d.metrado)} × ${cant(d.unitario, 6)}`).join('\n'))}">${cant(x.cantidad, 4)}</td><td class="n">${cant(x.redondeado, 4)}</td>
            <td><input class="n ${x.editado ? 'editado' : ''}" data-final value="${cant(x.final, 4)}" style="width:96px"></td><td class="n">${fmt(x.precio)}</td><td class="n">${fmt(r2(x.final * x.precio))}</td><td style="font-size:11.5px">${esc(x.alm) || (x.tipo === 'MAT' ? '<span class="tenue">sin enlace</span>' : '')}</td></tr>`).join('')}
          </tbody><tfoot><tr class="tot"><td colspan="7">TOTAL</td><td class="n">${fmt(total)}</td><td></td></tr></tfoot></table>` : '<div class="vacio">Sin resultados todavía.</div>'}</div></div>
    </div>`;

  enlazarExport(cont, { calc: () => rCalculadora(c.lineas, calcular(), { subtitulo: c.lineas.map((l) => `${l.item}: ${cant(l.metrado)} ${m.porItem.get(l.item)?.und || ''}`).join(' · ').slice(0, 170) }) });
  cont.querySelector('#b-add').onclick = async () => {
    const its = await elegirPartidas({ multiple: true, titulo: 'Agregar partidas a la calculadora' });
    its.forEach((it) => { if (!c.lineas.some((l) => l.item === it)) c.lineas.push({ item: it, metrado: 1 }); });
    if (its.length) { cambio(); refrescar(); }
  };
  cont.querySelector('#b-limpiar')?.addEventListener('click', () => { c.lineas = []; c.ajustes = {}; cambio(); refrescar(); });
  cont.querySelector('#b-reset')?.addEventListener('click', () => { c.ajustes = {}; cambio(); refrescar(); });
  cont.querySelector('#o-modo').onchange = (e) => { c.modo = e.target.value; cambio(); refrescar(); };
  cont.querySelector('#o-paso').onchange = (e) => { c.paso = Number(e.target.value); cambio(); refrescar(); };
  cont.querySelector('#o-mo').onchange = (e) => { c.mo = e.target.checked; cambio(); refrescar(); };
  cont.querySelector('#t-lin')?.addEventListener('change', (e) => {
    const el = e.target.closest('[data-k]');
    if (!el) return;
    c.lineas[Number(el.closest('tr').dataset.i)].metrado = num(el.value);
    cambio(); refrescar();
  });
  cont.querySelector('#t-lin')?.addEventListener('click', (e) => {
    const b = e.target.closest('[data-del]');
    if (!b) return;
    c.lineas.splice(Number(b.dataset.del), 1);
    cambio(); refrescar();
  });
  cont.querySelector('#t-ins')?.addEventListener('change', (e) => {
    const el = e.target.closest('[data-final]');
    if (!el) return;
    const k = el.closest('tr').dataset.k;
    const x = filas.find((f) => f.k === k);
    c.ajustes[k] = { teorico: x.cantidad, final: num(el.value) };
    cambio(); refrescar();
  });
  cont.querySelector('#b-salida')?.addEventListener('click', () => registrarSalida(filas));
  cont.querySelector('#b-met')?.addEventListener('click', () => registrarMetrado());
}

function registrarMetrado() {
  const c = estado();
  const o = S.obra;
  const campos = [{ id: 'fecha', et: 'Fecha', tipo: 'date', v: hoy(), req: true }, { id: 'modo', et: 'Si ya hay metrado ese día', tipo: 'select', ops: [['sumar', 'Sumar'], ['reemplazar', 'Reemplazar']] }];
  ventana({
    titulo: 'Registrar como metrado del día', ancho: '520px',
    cuerpo: htmlCampos(campos) + `<p class="tenue" style="font-size:12px">${c.lineas.length} partida(s) pasarán a los metrados diarios.</p>`,
    botones: [{ txt: 'Cancelar' }, { txt: 'Registrar', clase: 'btn-primario', accion: (cerrar, raiz) => {
      const v = leerCampos(raiz, campos);
      o.metrados[v.fecha] = o.metrados[v.fecha] || {};
      c.lineas.forEach((l) => { const prev = Number(o.metrados[v.fecha][l.item]) || 0; o.metrados[v.fecha][l.item] = r2((v.modo === 'sumar' ? prev : 0) + Number(l.metrado)); });
      cambio();
      aviso('Metrados registrados el ' + v.fecha.split('-').reverse().join('/'));
      ir('metrados', { dia: v.fecha });
    } }],
  });
}

function registrarSalida(filas) {
  const c = estado();
  const o = S.obra, m = S.motor;
  const mats = filas.filter((x) => x.tipo === 'MAT' && x.final > 0);
  if (!mats.length) { aviso('No hay materiales para la salida.', 'alerta'); return; }
  const stock = m.stock();
  const modo = c.modo === 'ninguno' || c.modo === 'dos' ? 'arriba' : c.modo;
  const lineas = mats.map((x) => ({ x, a: x.almItem, q: x.almItem ? x.qAlm : redondearCantidad(x.final, modo), crear: !x.almItem }));
  const destino = c.lineas.length === 1 ? c.lineas[0].item : '';
  ventana({
    titulo: 'Registrar salida de almacén', ancho: '960px',
    cuerpo: `${htmlCampos([
      { id: 'fecha', et: 'Fecha', tipo: 'date', v: hoy(), req: true }, { id: 'vale', et: 'N° de vale / PECOSA' }, { id: 'pedido_por', et: 'Pedido por', ph: 'MAESTRO DE OBRA' },
      { id: 'partida', et: 'Partida / destino', v: destino, ph: c.lineas.map((l) => l.item).join(', ').slice(0, 60) }, { id: 'observacion', et: 'Observación', col: 2, v: 'Calculado con A.C.U. × metrado' },
    ])}
      <div class="tabla-cont baja" style="margin-top:12px"><table class="t"><thead><tr><th><input type="checkbox" id="s-all" checked></th><th>Insumo (A.C.U.)</th><th>Ítem del almacén</th><th class="n">Saldo</th><th class="n">Cantidad a entregar</th><th>Und</th></tr></thead><tbody>
      ${lineas.map((l, i) => { const sal = l.a ? stock.get(l.a.codigo)?.saldo ?? 0 : 0; return `<tr data-i="${i}"><td><input type="checkbox" data-ok checked></td><td>${esc(l.x.insumo)}<div class="tenue" style="font-size:11px">${cant(l.x.final, 4)} ${esc(l.x.und)}</div></td>
        <td>${l.a ? `<span class="mono">${esc(l.a.codigo)}</span> ${esc(l.a.descripcion)}` : '<span class="badge badge-alerta">se creará en el catálogo del almacén</span>'}</td>
        <td class="n ${l.a && sal < l.q ? 'neg' : ''}">${l.a ? cant(sal, 2) : '—'}</td><td><input class="n" data-q value="${cant(l.q, 4)}" style="width:96px"></td><td>${esc(l.a ? l.a.und : l.x.und)}</td></tr>`; }).join('')}
      </tbody></table></div><p class="tenue" style="font-size:12px">Las cantidades del almacén usan la unidad y el factor de cada ítem (p. ej. varillas). Un saldo en rojo indica que la salida dejaría el almacén en negativo.</p>`,
    alAbrir: (raiz) => raiz.querySelector('#s-all').addEventListener('change', (e) => raiz.querySelectorAll('[data-ok]').forEach((x) => (x.checked = e.target.checked))),
    botones: [{ txt: 'Cancelar' }, { txt: 'Registrar salida', clase: 'btn-primario', icono: 'guardar', accion: (cerrar, raiz) => {
      const v = leerCampos(raiz, [{ id: 'fecha', et: 'Fecha', req: true }, { id: 'vale' }, { id: 'pedido_por' }, { id: 'partida' }, { id: 'observacion' }]);
      let n = 0, nuevos = 0;
      raiz.querySelectorAll('tbody tr').forEach((tr) => {
        if (!tr.querySelector('[data-ok]').checked) return;
        const l = lineas[Number(tr.dataset.i)];
        const q = num(tr.querySelector('[data-q]').value);
        if (!q) return;
        let a = l.a;
        if (!a) {
          a = { codigo: siguienteCodigoAlmacen(o), descripcion: l.x.insumo, und: l.x.und, grupo: 'MATERIALES', pu: l.x.precio, insumo_acu: l.x.insumo, factor: 1, observacion: 'Creado desde la calculadora' };
          o.almacen.push(a);
          nuevos++;
        }
        o.movimientos.push({ id: uid('M'), fecha: v.fecha, tipo: 'SALIDA', documento: 'VALE DE SALIDA', numero: '', oc_os: '', procedencia: '', partida: v.partida, codigo: a.codigo, cantidad: q, pedido_por: v.pedido_por, vale: v.vale, observacion: v.observacion, estado: 'CONFIRMADO', origen: 'CALCULADORA' });
        n++;
      });
      if (!n) throw new Error('No hay cantidades para registrar.');
      cambio({ guardarYa: true });
      aviso(`Salida registrada: ${n} ítem(s)${nuevos ? `, ${nuevos} ítem(s) nuevo(s) en el almacén` : ''}.`);
      ir('almacen', { tab: 'movimientos' });
    } }],
  });
}
