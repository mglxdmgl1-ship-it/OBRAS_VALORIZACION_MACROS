// Almacén: catálogo y saldos, movimientos (ingresos / salidas), salidas desde A.C.U. × metrado,
// kardex / bincard y cuadros 4.10, 4.11 y 4.12 del informe.
import { S, cambio } from '../store.js';
import { esc, fmt, cant, num, normal, hoy, mesDe, nombreMes, fFecha, uid, redondearCantidad, r2 } from '../util.js';
import { icono, ventana, formulario, confirmar, aviso, elegirArchivo, trabajando, paginar, htmlPaginador } from '../ui.js';
import { leerLibro, importarMovimientos, siguienteCodigoAlmacen } from '../importar.js';
import { rKardex, rMovDiario, rIngresosEgresos, rSaldoValorizado, rMovimientos, rVale, rCatalogoAlmacen, rPlantilla } from '../reportes.js';
import { cab, btnExport, enlazarExport, mesesDeObra, listaPartidasDatalist, datalistInsumos } from './comun.js';
import { refrescar } from '../app.js';

let tab = 'stock', fMes = '', fTipo = '', fq = '', pag = 1, kCod = '', cMes = null;
let gen = { desde: null, hasta: null, modo: null, props: null };
const DOCS = ['GUÍA DE REMISIÓN', 'ORDEN DE COMPRA', 'ORDEN DE SERVICIO', 'NEA', 'PECOSA', 'ACTA DE PRÉSTAMO', 'VALE DE SALIDA', 'ACTA DE DEVOLUCIÓN', 'OTRO'];
const GRUPOS = ['MATERIALES', 'HERRAMIENTAS', 'COMBUSTIBLE', 'EPP', 'SERVICIOS', 'EQUIPOS', 'OTROS'];

export function render(cont, params) {
  if (params.tab) tab = params.tab;
  if (params.desde) { gen = { desde: params.desde, hasta: params.hasta || params.desde, modo: null, props: null }; tab = 'generar'; }
  const o = S.obra, m = S.motor;
  const neg = m.movimientos().filter((x) => x.negativo).length;
  cont.innerHTML = `
    ${cab('Almacén de obra', `${o.almacen.length} ítems · ${o.movimientos.length} movimientos${neg ? ` · <span class="neg">${neg} movimiento(s) dejan saldo negativo</span>` : ''}. Cada ítem del almacén puede enlazarse a un insumo del A.C.U. (con factor de conversión, p. ej. varilla → kg) para el control de «NO VALORIZAR» y para generar salidas desde los metrados.`,
      `<button class="btn btn-primario" data-nuevo="INGRESO">${icono('mas')}Ingreso</button><button class="btn" data-nuevo="SALIDA">${icono('mas')}Salida</button><button class="btn" id="b-imp">${icono('importar')}Importar Excel</button>`)}
    <div class="pestanas">
      ${[['stock', 'Catálogo y saldos'], ['movimientos', 'Movimientos'], ['generar', 'Salidas desde metrados (A.C.U.)'], ['kardex', 'Kardex / bincard'], ['cuadros', 'Cuadros 4.10 · 4.11 · 4.12']].map(([k, t]) => `<button data-tab="${k}" class="${tab === k ? 'activo' : ''}">${t}</button>`).join('')}
    </div><div id="cuerpo"></div>${listaPartidasDatalist()}${datalistInsumos('dl-ins-acu', ['MAT', 'EQ', 'SC'])}
    <datalist id="dl-alm">${o.almacen.map((a) => `<option value="${esc(a.codigo)}">${esc(a.descripcion)} (${esc(a.und)})</option>`).join('')}</datalist>`;
  cont.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { tab = b.dataset.tab; pag = 1; history.replaceState(null, '', '#/almacen'); refrescar(); }));
  cont.querySelectorAll('[data-nuevo]').forEach((b) => (b.onclick = () => editarMov(null, b.dataset.nuevo)));
  cont.querySelector('#b-imp').onclick = importar;
  const box = cont.querySelector('#cuerpo');
  ({ stock, movimientos, generar, kardex, cuadros })[tab](box);
}

// =================================================================== CATÁLOGO
function stock(box) {
  const o = S.obra, m = S.motor;
  const st = m.stock();
  const nq = normal(fq);
  const lista = o.almacen.map((a, i) => ({ a, i })).filter(({ a }) => !nq || normal(a.codigo + ' ' + a.descripcion + ' ' + (a.insumo_acu || '')).includes(nq));
  const valor = o.almacen.reduce((s, a) => s + r2((st.get(a.codigo)?.saldo || 0) * (Number(a.pu) || 0)), 0);
  box.innerHTML = `
    <div class="filtros"><input type="search" class="buscar" id="q" placeholder="Buscar ítem…" value="${esc(fq)}"><span class="badge badge-azul">Saldo valorizado S/ ${fmt(valor)}</span><span style="flex:1"></span>
      <button class="btn btn-sm" id="b-desde-acu">${icono('acu')}Crear ítems desde el A.C.U.</button><button class="btn btn-sm btn-primario" id="b-item">${icono('mas')}Nuevo ítem</button>${btnExport('cat')}</div>
    <div class="tabla-cont"><table class="t" id="t-cat"><thead><tr><th>Código</th><th>Descripción</th><th>Und</th><th>Grupo</th><th class="n">P.U. S/</th><th>Insumo del A.C.U. (enlace)</th><th class="n" title="Unidades del A.C.U. que equivale 1 unidad del almacén">Factor</th><th class="n">Ingresos</th><th class="n">Salidas</th><th class="n">Saldo</th><th class="n">Valorizado</th><th></th></tr></thead><tbody>
    ${lista.map(({ a, i }) => { const s = st.get(a.codigo) || { ingresos: 0, salidas: 0, saldo: 0 }; const ins = a.insumo_acu ? m.insumo(a.insumo_acu) : null; return `<tr data-i="${i}">
      <td><input data-k="codigo" value="${esc(a.codigo)}" style="width:70px" class="mono"></td><td><input data-k="descripcion" value="${esc(a.descripcion)}" style="min-width:250px"></td><td><input data-k="und" value="${esc(a.und)}" style="width:70px"></td>
      <td><select data-k="grupo" style="width:130px">${GRUPOS.map((g) => `<option ${a.grupo === g ? 'selected' : ''}>${g}</option>`).join('')}</select></td><td><input class="n" data-k="pu" value="${fmt(a.pu)}" style="width:80px"></td>
      <td><input data-k="insumo_acu" list="dl-ins-acu" value="${esc(a.insumo_acu || '')}" style="min-width:230px" class="${a.insumo_acu && !ins ? 'editado' : ''}" title="${a.insumo_acu && !ins ? 'No existe en el catálogo de insumos' : ins ? 'Und A.C.U.: ' + esc(ins.und) : ''}"></td>
      <td><input class="n" data-k="factor" value="${cant(a.factor ?? 1, 4)}" style="width:70px"></td>
      <td class="n">${cant(s.ingresos, 2)}</td><td class="n">${cant(s.salidas, 2)}</td><td class="n ${s.saldo < 0 ? 'neg' : ''}"><b>${cant(s.saldo, 2)}</b></td><td class="n">${fmt(r2(s.saldo * (Number(a.pu) || 0)))}</td>
      <td class="acc"><button class="btn-icono" data-kx="${esc(a.codigo)}" title="Kardex">${icono('reportes')}</button><button class="btn-icono peligro" data-del="${i}">${icono('borrar')}</button></td></tr>`; }).join('') || '<tr><td colspan="12" class="vacio">Sin ítems. Créelos a mano, desde el A.C.U., o al importar movimientos.</td></tr>'}
    </tbody></table></div>`;
  enlazarExport(box, { cat: () => rCatalogoAlmacen() });
  let t;
  box.querySelector('#q').oninput = (e) => { clearTimeout(t); t = setTimeout(() => { fq = e.target.value; refrescar().then(() => { const el = document.querySelector('#q'); el?.focus(); el?.setSelectionRange(el.value.length, el.value.length); }); }, 250); };
  box.querySelector('#b-item').onclick = () => { o.almacen.push({ codigo: siguienteCodigoAlmacen(o), descripcion: 'NUEVO ÍTEM', und: 'UND', grupo: 'MATERIALES', pu: 0, insumo_acu: '', factor: 1, observacion: '' }); cambio(); refrescar(); };
  box.querySelector('#b-desde-acu').onclick = crearDesdeACU;
  box.querySelector('#t-cat tbody').addEventListener('change', (e) => {
    const el = e.target.closest('[data-k]');
    if (!el) return;
    const a = o.almacen[Number(el.closest('tr').dataset.i)];
    const k = el.dataset.k;
    if (k === 'codigo') {
      const nuevo = el.value.trim();
      if (!nuevo || o.almacen.some((x) => x !== a && x.codigo === nuevo)) { aviso('Código vacío o repetido.', 'error'); el.value = a.codigo; return; }
      o.movimientos.forEach((mv) => { if (mv.codigo === a.codigo) mv.codigo = nuevo; });
      a.codigo = nuevo;
    } else if (k === 'pu' || k === 'factor') { a[k] = num(el.value) || (k === 'factor' ? 1 : 0); }
    else if (k === 'insumo_acu') { const ins = S.motor.insumo(el.value.trim()); a.insumo_acu = ins ? ins.descripcion : el.value.trim(); }
    else a[k] = el.value.trim();
    cambio();
    if (k === 'pu' || k === 'codigo') refrescar();
  });
  box.querySelectorAll('[data-kx]').forEach((b) => (b.onclick = () => { kCod = b.dataset.kx; tab = 'kardex'; refrescar(); }));
  box.querySelectorAll('[data-del]').forEach((b) => (b.onclick = async () => {
    const a = o.almacen[Number(b.dataset.del)];
    const n = o.movimientos.filter((x) => x.codigo === a.codigo).length;
    if (!(await confirmar(`¿Eliminar el ítem <b>${esc(a.codigo)} ${esc(a.descripcion)}</b>?${n ? `<br><b class="neg">Tiene ${n} movimiento(s), que también se eliminarán.</b>` : ''}`, { peligro: true }))) return;
    o.movimientos = o.movimientos.filter((x) => x.codigo !== a.codigo);
    o.almacen.splice(Number(b.dataset.del), 1);
    cambio(); refrescar();
  }));
}
function crearDesdeACU() {
  const o = S.obra, m = S.motor;
  const enl = new Set(o.almacen.filter((a) => a.insumo_acu).map((a) => normal(a.insumo_acu)));
  const usados = new Set();
  Object.values(o.acu).forEach((fs) => fs.forEach((f) => f.insumo && f.tipo === 'MAT' && usados.add(normal(f.insumo))));
  const cand = o.insumos.filter((i) => i.tipo === 'MAT' && usados.has(normal(i.descripcion)) && !enl.has(normal(i.descripcion)) && m.controlaInsumo({ tipo: 'MAT', descripcion: i.descripcion }));
  if (!cand.length) { aviso('Todos los materiales del A.C.U. ya tienen ítem en el almacén.'); return; }
  ventana({
    titulo: 'Crear ítems del almacén desde los materiales del A.C.U.', ancho: '820px',
    cuerpo: `<p>${cand.length} material(es) del A.C.U. no tienen ítem en el almacén. Se crearán con la misma unidad, precio y factor 1 (puede cambiarlos después).</p>
      <div class="tabla-cont baja"><table class="t"><thead><tr><th><input type="checkbox" id="c-all" checked></th><th>Código</th><th>Insumo</th><th>Und</th><th class="n">Precio</th></tr></thead><tbody>
      ${cand.map((i, k) => `<tr><td><input type="checkbox" data-k="${k}" checked></td><td class="mono tenue">${esc(i.codigo)}</td><td>${esc(i.descripcion)}</td><td>${esc(i.und)}</td><td class="n">${fmt(i.precio)}</td></tr>`).join('')}</tbody></table></div>`,
    alAbrir: (r) => r.querySelector('#c-all').addEventListener('change', (e) => r.querySelectorAll('[data-k]').forEach((c) => (c.checked = e.target.checked))),
    botones: [{ txt: 'Cancelar' }, { txt: 'Crear ítems', clase: 'btn-primario', accion: (c, r) => {
      let n = 0;
      r.querySelectorAll('[data-k]:checked').forEach((ch) => { const i = cand[Number(ch.dataset.k)]; o.almacen.push({ codigo: siguienteCodigoAlmacen(o), descripcion: i.descripcion, und: i.und, grupo: 'MATERIALES', pu: i.precio, insumo_acu: i.descripcion, factor: 1, observacion: 'Creado desde el A.C.U.' }); n++; });
      cambio(); aviso(`${n} ítem(s) creado(s).`); refrescar();
    } }],
  });
}

// =================================================================== MOVIMIENTOS
function movimientos(box) {
  const o = S.obra, m = S.motor;
  const nq = normal(fq);
  const todos = m.movimientos().slice().reverse();
  const lista = todos.filter((x) => (!fMes || mesDe(x.fecha) === fMes) && (!fTipo || x.tipo === fTipo) && (!nq || normal([x.codigo, m.itemAlmacen(x.codigo)?.descripcion, x.documento, x.numero, x.vale, x.partida, x.procedencia, x.pedido_por].join(' ')).includes(nq)));
  const pg = paginar(lista, pag, 100);
  box.innerHTML = `
    <div class="filtros">
      <select id="f-mes"><option value="">Todos los meses</option>${mesesDeObra().map((x) => `<option value="${x}" ${fMes === x ? 'selected' : ''}>${nombreMes(x)}</option>`).join('')}</select>
      <select id="f-tipo"><option value="">Todos</option>${['INGRESO', 'SALIDA', 'DEVOLUCION'].map((t) => `<option ${fTipo === t ? 'selected' : ''}>${t}</option>`).join('')}</select>
      <input type="search" class="buscar" id="q" placeholder="Buscar ítem, documento, vale, partida…" value="${esc(fq)}">
      <span style="flex:1"></span><button class="btn btn-sm" id="b-vale" disabled>${icono('pdf')}Vale de salida (marcados)</button><button class="btn btn-sm" id="b-plant">${icono('excel')}Plantilla</button>${btnExport('movs')}
    </div>
    <div class="tabla-cont"><table class="t" id="t-mov"><thead><tr><th><input type="checkbox" id="m-all"></th><th>Fecha</th><th>Tipo</th><th>Documento</th><th>Vale</th><th>Procedencia / pedido por</th><th>Partida</th><th>Ítem</th><th class="n">Cantidad</th><th class="n">Saldo después</th><th>Estado</th><th></th></tr></thead><tbody>
    ${pg.filas.map((x) => { const a = m.itemAlmacen(x.codigo); return `<tr data-id="${esc(x.id)}" class="${x.negativo ? 'e-novalorizar' : ''}"><td><input type="checkbox" data-sel></td><td>${fFecha(x.fecha)}</td><td><span class="badge ${x.tipo === 'INGRESO' ? 'badge-ok' : x.tipo === 'SALIDA' ? 'badge-alerta' : 'badge-azul'}">${x.tipo}</span></td>
      <td style="font-size:12px">${esc([x.documento, x.numero].filter(Boolean).join(' '))}${x.oc_os ? `<div class="tenue">${esc(x.oc_os)}</div>` : ''}</td><td>${esc(x.vale || '')}</td><td style="font-size:12px">${esc(x.tipo === 'SALIDA' ? x.pedido_por || '' : x.procedencia || '')}</td><td class="item">${esc(x.partida || '')}</td>
      <td style="font-size:12px"><span class="mono">${esc(x.codigo)}</span> ${esc(a?.descripcion || '¿ítem inexistente?')}</td><td class="n"><b>${cant(x.cantidad, 4)}</b> <span class="tenue">${esc(a?.und || '')}</span></td><td class="n ${x.negativo ? 'neg' : ''}">${cant(x.saldo, 2)}</td>
      <td>${x.estado === 'POR CONFIRMAR' ? '<span class="badge badge-alerta">POR CONFIRMAR</span>' : ''} <span class="tenue" style="font-size:11px">${esc(x.origen || '')}</span></td>
      <td class="acc"><button class="btn-icono" data-ed title="Editar">${icono('editar')}</button><button class="btn-icono peligro" data-del title="Eliminar">${icono('borrar')}</button></td></tr>`; }).join('') || '<tr><td colspan="12" class="vacio">Sin movimientos.</td></tr>'}
    </tbody></table></div>${htmlPaginador(pg.pagina, pg.paginas, lista.length)}`;
  enlazarExport(box, { movs: () => rMovimientos(lista.slice().reverse(), 'MOVIMIENTOS DEL ALMACÉN', fMes ? nombreMes(fMes) : '') });
  box.querySelector('#f-mes').onchange = (e) => { fMes = e.target.value; pag = 1; refrescar(); };
  box.querySelector('#f-tipo').onchange = (e) => { fTipo = e.target.value; pag = 1; refrescar(); };
  let t;
  box.querySelector('#q').oninput = (e) => { clearTimeout(t); t = setTimeout(() => { fq = e.target.value; pag = 1; refrescar().then(() => { const el = document.querySelector('#q'); el?.focus(); el?.setSelectionRange(el.value.length, el.value.length); }); }, 300); };
  box.querySelectorAll('[data-pag]').forEach((b) => (b.onclick = () => { pag = Number(b.dataset.pag); refrescar(); }));
  box.querySelector('#b-plant').onclick = async () => { const { aExcel } = await import('../exportar.js'); aExcel(rPlantilla('movimientos')); };
  const tb = box.querySelector('#t-mov tbody');
  const bVale = box.querySelector('#b-vale');
  const marcados = () => [...tb.querySelectorAll('[data-sel]:checked')].map((c) => o.movimientos.find((x) => x.id === c.closest('tr').dataset.id)).filter(Boolean);
  tb.addEventListener('change', () => (bVale.disabled = !marcados().some((x) => x.tipo === 'SALIDA')));
  box.querySelector('#m-all').onchange = (e) => { tb.querySelectorAll('[data-sel]').forEach((c) => (c.checked = e.target.checked)); bVale.disabled = !marcados().some((x) => x.tipo === 'SALIDA'); };
  bVale.onclick = async () => { const { aPDF } = await import('../exportar.js'); aPDF(rVale(marcados().filter((x) => x.tipo === 'SALIDA'))); };
  tb.addEventListener('click', async (e) => {
    const tr = e.target.closest('tr[data-id]');
    if (!tr) return;
    const id = tr.dataset.id;
    if (e.target.closest('[data-ed]')) editarMov(id);
    if (e.target.closest('[data-del]')) {
      if (!(await confirmar('¿Eliminar este movimiento?', { peligro: true }))) return;
      o.movimientos = o.movimientos.filter((x) => x.id !== id);
      cambio(); refrescar();
    }
  });
}
export function editarMov(id, tipo = 'INGRESO') {
  const o = S.obra;
  const mv = id ? o.movimientos.find((x) => x.id === id) : null;
  const val = mv ? { ...mv } : { fecha: hoy(), tipo, estado: 'CONFIRMADO', documento: tipo === 'SALIDA' ? 'VALE DE SALIDA' : 'GUÍA DE REMISIÓN' };
  const campos = [
    { id: 'fecha', et: 'Fecha', tipo: 'date', req: true }, { id: 'tipo', et: 'Tipo', tipo: 'select', ops: ['INGRESO', 'SALIDA', 'DEVOLUCION'] },
    { id: 'estado', et: 'Estado', tipo: 'select', ops: ['CONFIRMADO', 'POR CONFIRMAR'], ayuda: 'POR CONFIRMAR: guía observada, sin NEA/PECOSA…' },
    { id: 'codigo', et: 'Ítem del almacén (código)', req: true, lista: 'dl-alm', col: 2, ayuda: 'Escriba el código o elija de la lista' }, { id: 'cantidad', et: 'Cantidad', tipo: 'number', req: true },
    { id: 'documento', et: 'Documento', tipo: 'select', ops: DOCS }, { id: 'numero', et: 'N° de documento' }, { id: 'oc_os', et: 'O/C – O/S' },
    { id: 'procedencia', et: 'Procedencia / proveedor', col: 2 }, { id: 'vale', et: 'N° vale / PECOSA' },
    { id: 'partida', et: 'Partida (destino)', lista: 'dl-partidas' }, { id: 'pedido_por', et: 'Pedido / recibido por' }, { id: 'partida_cubre', et: 'Partida que cubre (opcional)', lista: 'dl-partidas', ayuda: 'Si el documento ampara todos los insumos de una partida (p. ej. EPP)' },
    { id: 'observacion', et: 'Observación', col: 3 },
  ];
  formulario({
    titulo: mv ? 'Editar movimiento' : 'Nuevo ' + tipo.toLowerCase(), campos, val, ancho: '860px',
    alGuardar: (v) => {
      v.codigo = v.codigo.split(' ')[0].trim();
      if (!o.almacen.some((a) => a.codigo === v.codigo)) throw new Error('No existe el ítem ' + v.codigo + ' en el catálogo del almacén. Créelo primero en «Catálogo y saldos».');
      if (!(v.cantidad > 0)) throw new Error('La cantidad debe ser mayor que cero.');
      if (mv) Object.assign(mv, v);
      else o.movimientos.push({ id: uid('M'), origen: 'DIGITADO', ...v });
      cambio();
      const sal = S.motor.saldoActual(v.codigo);
      aviso(`Guardado. Saldo de ${v.codigo}: ${cant(sal, 2)}`, sal < 0 ? 'alerta' : 'ok');
      refrescar();
    },
  });
}

// =================================================================== GENERAR SALIDAS
function generar(box) {
  const o = S.obra, m = S.motor;
  const fs = m.fechasConMetrado();
  if (!gen.desde) { gen.desde = fs.length ? fs[fs.length - 1] : hoy(); gen.hasta = gen.desde; }
  if (!gen.modo) gen.modo = o.parametros.redondeo === 'ninguno' || o.parametros.redondeo === 'dos' ? 'arriba' : o.parametros.redondeo || 'arriba';
  const { propuestas, sinEnlace } = m.proponerSalidas(gen.desde, gen.hasta, gen.modo);
  const existentes = new Set(o.movimientos.filter((x) => x.tipo === 'SALIDA').map((x) => `${x.fecha}|${x.partida}|${x.codigo}`));
  if (!gen.props || gen.key !== `${gen.desde}|${gen.hasta}|${gen.modo}|${propuestas.length}`) {
    gen.props = propuestas.map((p) => ({ ...p, ok: !existentes.has(`${p.fecha}|${p.partida}|${p.codigo}`) && p.cantidad > 0, dup: existentes.has(`${p.fecha}|${p.partida}|${p.codigo}`) }));
    gen.key = `${gen.desde}|${gen.hasta}|${gen.modo}|${propuestas.length}`;
  }
  const P = gen.props;
  box.innerHTML = `
    <div class="nota">Calcula los materiales que salen del almacén con el A.C.U. × el metrado de cada día (como «ImportarControl» del programa de almacén). Las cantidades se convierten a la unidad del almacén con el factor y se redondean; puede <b>editarlas</b> antes de registrar.</div>
    <div class="filtros"><label>Desde <input type="date" id="g-d" value="${gen.desde}"></label><label>Hasta <input type="date" id="g-h" value="${gen.hasta}"></label>
      <label>Redondeo <select id="g-m">${[['arriba', 'Entero hacia arriba'], ['normal', 'Entero más cercano'], ['abajo', 'Entero hacia abajo'], ['dos', '2 decimales'], ['ninguno', 'Sin redondeo']].map(([k, t]) => `<option value="${k}" ${gen.modo === k ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
      <span style="flex:1"></span><button class="btn btn-primario" id="g-ok" ${P.length ? '' : 'disabled'}>${icono('guardar')}Registrar salidas marcadas</button></div>
    ${sinEnlace.length ? `<div class="nota nota-alerta"><b>${sinEnlace.length} material(es) del A.C.U. no tienen ítem en el almacén</b> (no se pueden generar sus salidas): ${esc(sinEnlace.slice(0, 8).map((x) => `${x.insumo} (${cant(x.cantidad, 2)})`).join(' · '))}${sinEnlace.length > 8 ? '…' : ''}. <button class="btn btn-sm" id="g-crear">Crear y enlazar ítems</button></div>` : ''}
    ${P.length ? `<div class="tabla-cont"><table class="t" id="t-gen"><thead><tr><th><input type="checkbox" id="g-all"></th><th>Fecha</th><th>Partida</th><th class="n">Metrado</th><th>Insumo del A.C.U.</th><th>Ítem del almacén</th><th class="n">Según A.C.U.</th><th class="n">Cantidad a registrar</th><th>Und</th><th class="n">Saldo actual</th><th></th></tr></thead><tbody>
      ${P.map((p, i) => `<tr data-i="${i}" class="${p.dup ? 'tenue' : ''}"><td><input type="checkbox" data-ok ${p.ok ? 'checked' : ''}></td><td>${fFecha(p.fecha)}</td><td class="item">${esc(p.partida)}</td><td class="n">${cant(p.metrado, 2)}</td><td style="font-size:12px">${esc(p.insumo_acu)}</td><td style="font-size:12px"><span class="mono">${esc(p.codigo)}</span> ${esc(p.descripcion)}</td><td class="n tenue">${cant(p.teorico, 4)}</td>
        <td><input class="n" data-q value="${cant(p.cantidad, 4)}" style="width:90px"></td><td>${esc(p.und)}</td><td class="n ${(m.stock().get(p.codigo)?.saldo || 0) < p.cantidad ? 'neg' : ''}">${cant(m.stock().get(p.codigo)?.saldo || 0, 2)}</td><td>${p.dup ? '<span class="badge badge-alerta">ya registrada</span>' : ''}</td></tr>`).join('')}
    </tbody></table></div>` : `<div class="panel"><div class="vacio"><h3>Nada que generar</h3><p>No hay metrados en el rango con materiales enlazados al almacén.</p></div></div>`}`;
  const recalc = () => { gen.props = null; refrescar(); };
  box.querySelector('#g-d').onchange = (e) => { gen.desde = e.target.value; if (gen.hasta < gen.desde) gen.hasta = gen.desde; recalc(); };
  box.querySelector('#g-h').onchange = (e) => { gen.hasta = e.target.value; recalc(); };
  box.querySelector('#g-m').onchange = (e) => { gen.modo = e.target.value; recalc(); };
  box.querySelector('#g-crear')?.addEventListener('click', () => {
    let n = 0;
    for (const s of sinEnlace) { const i = m.insumo(s.insumo); o.almacen.push({ codigo: siguienteCodigoAlmacen(o), descripcion: s.insumo, und: i?.und || 'UND', grupo: 'MATERIALES', pu: i?.precio || 0, insumo_acu: s.insumo, factor: 1, observacion: 'Creado al generar salidas' }); n++; }
    cambio(); aviso(`${n} ítem(s) creado(s) y enlazado(s).`); recalc();
  });
  const tb = box.querySelector('#t-gen tbody');
  tb?.addEventListener('change', (e) => {
    const tr = e.target.closest('tr[data-i]');
    if (!tr) return;
    const p = P[Number(tr.dataset.i)];
    if (e.target.matches('[data-ok]')) p.ok = e.target.checked;
    if (e.target.matches('[data-q]')) { p.cantidad = num(e.target.value); e.target.classList.add('editado'); }
  });
  box.querySelector('#g-all')?.addEventListener('change', (e) => { P.forEach((p) => (p.ok = e.target.checked)); tb.querySelectorAll('[data-ok]').forEach((c) => (c.checked = e.target.checked)); });
  box.querySelector('#g-ok')?.addEventListener('click', async () => {
    const sel = P.filter((p) => p.ok && p.cantidad > 0);
    if (!sel.length) { aviso('No hay filas marcadas con cantidad.', 'alerta'); return; }
    if (!(await confirmar(`¿Registrar ${sel.length} salida(s) del almacén?`))) return;
    sel.forEach((p) => o.movimientos.push({ id: uid('M'), fecha: p.fecha, tipo: 'SALIDA', documento: 'VALE DE SALIDA', numero: '', oc_os: '', procedencia: '', partida: p.partida, codigo: p.codigo, cantidad: p.cantidad, pedido_por: '', vale: '', observacion: `A.C.U. × metrado (${cant(p.metrado, 2)}); teórico ${cant(p.teorico, 4)}`, estado: 'CONFIRMADO', origen: 'A.C.U.' }));
    cambio({ guardarYa: true });
    gen.props = null;
    aviso(`${sel.length} salida(s) registrada(s).`);
    tab = 'movimientos'; refrescar();
  });
}

// =================================================================== KARDEX
function kardex(box) {
  const o = S.obra, m = S.motor;
  if (!kCod || !o.almacen.some((a) => a.codigo === kCod)) kCod = o.almacen[0]?.codigo || '';
  const k = kCod ? m.kardex(kCod) : null;
  box.innerHTML = `
    <div class="filtros"><select id="k-cod" style="min-width:380px">${o.almacen.map((a) => `<option value="${esc(a.codigo)}" ${a.codigo === kCod ? 'selected' : ''}>${esc(a.codigo)} · ${esc(a.descripcion)}</option>`).join('')}</select>
      <span style="flex:1"></span>${btnExport('k1', 'Kardex')}${btnExport('b1', 'Bincard')}${btnExport('kall', 'Todos')}</div>
    ${k ? `<div class="tabla-cont"><table class="t"><thead><tr><th rowspan="2">Fecha</th><th rowspan="2">Documento</th><th rowspan="2">Procedencia / destino</th><th colspan="2" class="c">Entrada</th><th colspan="2" class="c">Salida</th><th colspan="2" class="c">Saldo</th></tr><tr><th class="n">Cant.</th><th class="n">S/</th><th class="n">Cant.</th><th class="n">S/</th><th class="n">Cant.</th><th class="n">S/</th></tr></thead><tbody>
      ${k.filas.map((f) => `<tr class="${f.negativo ? 'e-novalorizar' : ''}"><td>${fFecha(f.fecha)}</td><td style="font-size:12px">${esc([f.tipo, f.documento, f.numero, f.vale && 'VALE ' + f.vale].filter(Boolean).join(' '))}</td><td style="font-size:12px">${esc(f.tipo === 'SALIDA' ? [f.partida, f.pedido_por].filter(Boolean).join(' · ') : f.procedencia || '')}</td><td class="n">${cant(f.entrada, 2) || ''}</td><td class="n">${f.entradaS ? fmt(f.entradaS) : ''}</td><td class="n">${f.salida ? cant(f.salida, 2) : ''}</td><td class="n">${f.salidaS ? fmt(f.salidaS) : ''}</td><td class="n"><b>${cant(f.saldo, 2)}</b></td><td class="n">${fmt(f.saldoS)}</td></tr>`).join('') || '<tr><td colspan="9" class="vacio">Sin movimientos para este ítem.</td></tr>'}
    </tbody></table></div><p class="tenue" style="font-size:12px">P.U. del ítem: S/ ${fmt(k.pu)} · ${esc(k.item?.und || '')}${k.item?.insumo_acu ? ' · A.C.U.: ' + esc(k.item.insumo_acu) + ' (factor ' + cant(k.item.factor, 4) + ')' : ''}</p>` : '<div class="vacio">Sin ítems en el almacén.</div>'}`;
  enlazarExport(box, { k1: () => rKardex([kCod]), b1: () => rKardex([kCod], null, null, true), kall: () => rKardex(null) });
  box.querySelector('#k-cod')?.addEventListener('change', (e) => { kCod = e.target.value; refrescar(); });
}

// =================================================================== CUADROS
function cuadros(box) {
  const m = S.motor;
  const meses = mesesDeObra();
  if (!cMes) cMes = meses.filter((x) => S.obra.movimientos.some((mv) => mv.fecha.startsWith(x))).pop() || meses[meses.length - 1] || mesDe(hoy());
  const res = m.resumenAlmacen(cMes).filter((r) => r.anterior || r.ingresos || r.salidas);
  const md = m.movimientoDiario(cMes);
  box.innerHTML = `
    <div class="filtros"><select id="c-mes">${meses.map((x) => `<option value="${x}" ${x === cMes ? 'selected' : ''}>${nombreMes(x)}</option>`).join('')}</select><span class="tenue">${res.length} ítem(s) con movimiento o saldo · ${md.dias.length} día(s) con movimiento</span></div>
    <div class="rep-lista" style="margin-bottom:14px">
      <div class="rep"><h4>4.10 Movimiento diario de almacén</h4><p>Saldo anterior y entradas / salidas de cada día del mes.</p><div class="acciones">${btnExport('c410')}</div></div>
      <div class="rep"><h4>4.11 Ingresos – egresos</h4><p>Saldo anterior, ingresos y egresos del mes y saldo actual (con fórmulas).</p><div class="acciones">${btnExport('c411')}</div></div>
      <div class="rep"><h4>4.12 Saldo valorizado</h4><p>Saldo × P.U. de cada ítem al cierre del mes.</p><div class="acciones">${btnExport('c412')}</div></div>
    </div>
    <div class="tabla-cont"><table class="t"><thead><tr><th>Código</th><th>Descripción</th><th>Und</th><th class="n">Saldo anterior</th><th class="n">Ingresos</th><th class="n">Egresos</th><th class="n">Saldo</th><th class="n">P.U.</th><th class="n">Valorizado S/</th></tr></thead><tbody>
      ${res.map((r) => `<tr><td class="mono">${esc(r.codigo)}</td><td>${esc(r.descripcion)}</td><td>${esc(r.und)}</td><td class="n">${cant(r.anterior, 2)}</td><td class="n">${cant(r.ingresos, 2)}</td><td class="n">${cant(r.salidas, 2)}</td><td class="n ${r.saldo < 0 ? 'neg' : ''}"><b>${cant(r.saldo, 2)}</b></td><td class="n">${fmt(r.pu)}</td><td class="n">${fmt(r.valorizado)}</td></tr>`).join('') || '<tr><td colspan="9" class="vacio">Sin movimientos.</td></tr>'}
    </tbody><tfoot><tr class="tot"><td colspan="8">TOTAL SALDO VALORIZADO</td><td class="n">${fmt(res.reduce((s, r) => s + r.valorizado, 0))}</td></tr></tfoot></table></div>`;
  enlazarExport(box, { c410: () => rMovDiario(cMes), c411: () => rIngresosEgresos(cMes), c412: () => rSaldoValorizado(cMes) });
  box.querySelector('#c-mes').onchange = (e) => { cMes = e.target.value; refrescar(); };
}

async function importar() {
  const f = await elegirArchivo('.xlsx,.xlsm,.csv');
  if (!f) return;
  try {
    const hojas = await trabajando('Leyendo…', () => leerLibro(f));
    const r = importarMovimientos(S.obra, hojas[0].filas);
    cambio({ guardarYa: true });
    aviso(`${r.n} movimiento(s) importado(s), ${r.nuevos} ítem(s) nuevo(s).${r.errores.length ? ' Errores: ' + r.errores.slice(0, 3).join('; ') : ''}`, r.errores.length ? 'alerta' : 'ok', 9000);
    tab = 'movimientos'; refrescar();
  } catch (e) { aviso(e.message, 'error', 8000); }
}
export { redondearCantidad };
