// Análisis de costos unitarios: editar insumos, cuadrilla, cantidades y precios; importar desde Excel.
import { S, cambio } from '../store.js';
import { esc, fmt, cant, num, normal, r2 } from '../util.js';
import { icono, ventana, confirmar, aviso, elegirArchivo, trabajando } from '../ui.js';
import { leerLibro, analizarACU, aplicarACU } from '../importar.js';
import { rACU, rPlantilla } from '../reportes.js';
import { TIPOS_INSUMO } from '../calc.js';
import { cab, btnExport, enlazarExport, elegirPartidas, datalistInsumos, sinPresupuesto } from './comun.js';
import { refrescar, ir } from '../app.js';

let sel = null;
let q = '';
let soloSin = false;

export function render(cont, params) {
  const o = S.obra, m = S.motor;
  if (!m.partidasP.length) { cont.innerHTML = cab('Análisis de costos unitarios (A.C.U.)', '') + sinPresupuesto(); return; }
  if (params.item && m.porItem.get(params.item)?.tipo === 'P') sel = params.item;
  if (!sel || !m.porItem.get(sel)) sel = m.partidasP[0].item;
  const sinACU = m.partidasP.filter((p) => !m.acu(p.item).length).length;
  cont.innerHTML = `
    ${cab('Análisis de costos unitarios (A.C.U.)', `${m.partidasP.length - sinACU} de ${m.partidasP.length} partidas con análisis · ${o.insumos.length} insumos en el catálogo. Al cambiar el precio de un insumo se actualizan <b>todos</b> los A.C.U. que lo usan.`,
      `<button class="btn btn-oro" id="b-imp">${icono('importar')}Importar A.C.U. (Excel)</button>${btnExport('todos', 'Todos')}`)}
    <div class="rejilla-3" style="grid-template-columns: minmax(280px, 1fr) 2.4fr">
      <div class="panel" style="margin:0"><div class="panel-cab"><input type="search" id="a-q" placeholder="Buscar partida…" value="${esc(q)}" style="flex:1"><label class="chk" style="font-size:12px"><input type="checkbox" id="a-sin" data-libre ${soloSin ? 'checked' : ''}> sin A.C.U.</label></div>
        <div class="tabla-cont" style="border:0;border-radius:0;max-height:calc(100vh - 250px)"><table class="t"><tbody id="a-lista"></tbody></table></div></div>
      <div id="a-det"></div>
    </div>${datalistInsumos()}`;
  enlazarExport(cont, { todos: () => rACU(null) });
  cont.querySelector('#b-imp').onclick = importarACU;
  const pintaLista = () => {
    const nq = normal(q);
    const lista = m.partidasP.filter((p) => (!nq || normal(p.item + ' ' + p.descripcion).includes(nq)) && (!soloSin || !m.acu(p.item).length));
    cont.querySelector('#a-lista').innerHTML = lista.slice(0, 700).map((p) => `<tr data-it="${esc(p.item)}" class="${p.item === sel ? 'sel' : ''}" style="cursor:pointer"><td class="item">${esc(p.item)}</td><td><span class="desc-trunc" style="max-width:260px">${esc(p.descripcion)}</span></td><td class="n">${m.acu(p.item).length ? '' : '<span class="badge badge-alerta">sin</span>'}</td></tr>`).join('') || '<tr><td class="tenue">Sin resultados</td></tr>';
  };
  pintaLista();
  cont.querySelector('#a-q').oninput = (e) => { q = e.target.value; pintaLista(); };
  cont.querySelector('#a-sin').onchange = (e) => { soloSin = e.target.checked; pintaLista(); };
  cont.querySelector('#a-lista').onclick = (e) => {
    const tr = e.target.closest('tr[data-it]');
    if (!tr) return;
    sel = tr.dataset.it;
    cont.querySelectorAll('#a-lista tr').forEach((x) => x.classList.toggle('sel', x.dataset.it === sel));
    detalle(cont.querySelector('#a-det'));
  };
  detalle(cont.querySelector('#a-det'));
  setTimeout(() => cont.querySelector('#a-lista tr.sel')?.scrollIntoView({ block: 'center' }), 30);
}

const ORD = { MO: 0, MAT: 1, EQ: 2, SC: 3, '%MO': 4, SP: 5 };
function detalle(box) {
  const o = S.obra;
  const item = sel;
  const p = S.motor.porItem.get(item);
  const filas = o.acu[item] || (o.acu[item] = []);
  const m = S.motor;
  const calc = m.acu(item);
  const res = m.resumenACU(item);
  const puA = m.puACU(item);
  const jor = Number(o.parametros.jornada) || 8;
  box.innerHTML = `<div class="panel" style="margin:0">
    <div class="panel-cab"><div><h2><span class="mono">${esc(item)}</span> · ${esc(p.descripcion)}</h2>
      <div class="tenue" style="font-size:12px;margin-top:3px">Unidad <b>${esc(p.und)}</b> · Metrado E.T. <b>${fmt(p.metrado)}</b> · Parcial S/ ${fmt(m.parcial(p))}</div></div>
      <div class="acciones">${btnExport('uno')}</div></div>
    <div class="panel-cuerpo">
      <div class="tarjetas" style="grid-template-columns:repeat(auto-fill,minmax(150px,1fr));margin-bottom:12px">
        <div class="tarjeta"><div class="et">P.U. presupuesto</div><div class="val">${fmt(p.pu)}</div></div>
        <div class="tarjeta"><div class="et">P.U. según A.C.U.</div><div class="val ${Math.abs(puA - p.pu) >= 0.01 ? 'neg' : ''}" id="pu-acu">${fmt(puA)}</div><div class="sub">${Math.abs(puA - p.pu) >= 0.01 ? 'Difiere del presupuesto' : 'Coincide'}</div></div>
        <div class="tarjeta"><div class="et">Rendimiento</div><div style="display:flex;gap:6px;align-items:center;margin-top:4px"><input id="rend" value="${p.rendimiento ?? ''}" style="width:90px;text-align:right"><span class="tenue">${esc(p.und)}/día</span></div><div class="sub">Jornada ${jor} h</div></div>
        <div class="tarjeta"><div class="et">M.O. / Mat. / Eq.</div><div class="sub" style="font-size:12.5px;color:var(--texto);margin-top:6px">${fmt(res.MO + res['%MO'])} / ${fmt(res.MAT)} / ${fmt(res.EQ + res.SC)}</div></div>
      </div>
      <div class="acciones" style="margin-bottom:10px">
        <button class="btn btn-sm btn-primario" id="b-add">${icono('mas')}Agregar insumo</button>
        <button class="btn btn-sm" id="b-copiar">${icono('copiar')}Copiar A.C.U. de otra partida</button>
        <button class="btn btn-sm" id="b-cuad" title="cantidad = cuadrilla × jornada / rendimiento">Recalcular con cuadrilla</button>
        <button class="btn btn-sm" id="b-usarpu" ${Math.abs(puA - p.pu) < 0.01 ? 'disabled' : ''}>Usar este P.U. en el presupuesto</button>
        <button class="btn btn-sm" id="b-calc">${icono('calc')}Calcular insumos para un metrado</button>
        <button class="btn btn-sm btn-peligro" id="b-vaciar" ${filas.length ? '' : 'disabled'}>${icono('borrar')}Vaciar</button>
      </div>
      <div class="tabla-cont libre"><table class="t" id="t-acu"><thead><tr><th>Tipo</th><th>Código</th><th>Insumo / recurso</th><th>Und</th><th class="n">Cuadrilla</th><th class="n">Cantidad</th><th class="n">Precio S/</th><th class="n">Parcial S/</th><th class="n">% P.U.</th><th></th></tr></thead><tbody>
      ${filas.map((f, i) => ({ f, i, c: calc[i] })).sort((a, b) => (ORD[a.f.tipo] ?? 9) - (ORD[b.f.tipo] ?? 9) || a.i - b.i).map(({ f, i, c }) => `<tr data-r="${i}">
        <td><select data-k="tipo" style="width:74px">${Object.keys(TIPOS_INSUMO).map((t) => `<option value="${t}" ${f.tipo === t ? 'selected' : ''}>${t}</option>`).join('')}</select></td>
        <td class="mono tenue" style="font-size:11px">${esc(c.codigo || '')}</td>
        <td style="min-width:280px">${f.tipo === '%MO' ? `<input data-k="descripcion" value="${esc(f.descripcion || 'HERRAMIENTAS MANUALES')}">` : `<input data-k="insumo" list="dl-insumos" value="${esc(f.insumo || '')}" class="${c.existe ? '' : 'editado'}" title="${c.existe ? '' : 'Insumo nuevo: se agregará al catálogo'}">`}</td>
        <td>${f.tipo === '%MO' ? '%mo' : `<input data-k="und" value="${esc(c.und || '')}" style="width:58px">`}</td>
        <td><input class="n" data-k="cuadrilla" value="${f.cuadrilla ?? ''}" style="width:72px" ${f.tipo === 'MO' || f.tipo === 'EQ' ? '' : 'disabled'}></td>
        <td><input class="n" data-k="cantidad" value="${cant(f.cantidad, 6)}" style="width:96px" title="${f.tipo === '%MO' ? 'Porcentaje de la mano de obra' : 'Cantidad por ' + esc(p.und)}"></td>
        <td>${f.tipo === '%MO' ? `<span class="n" style="display:block;text-align:right" data-c="precio">${fmt(c.precio, 4)}</span>` : `<input class="n" data-k="precio" value="${fmt(c.precio, 2)}" style="width:86px" title="Cambia el precio del insumo en todos los A.C.U.">`}</td>
        <td class="n" data-c="parcial">${fmt(c.parcial, 4)}</td>
        <td class="n tenue" data-c="pct">${p.pu ? fmt((c.parcial / p.pu) * 100, 1) + ' %' : ''}</td>
        <td><button class="btn-icono peligro" data-del="${i}" title="Quitar">${icono('borrar')}</button></td></tr>`).join('') || '<tr><td colspan="10" class="vacio">Esta partida no tiene A.C.U. Agregue los insumos o cópielos de otra partida.</td></tr>'}
      </tbody><tfoot><tr class="tot"><td colspan="7">COSTO UNITARIO DIRECTO POR ${esc((p.und || '').toUpperCase())}</td><td class="n" id="tot-acu">${fmt(puA, 4)}</td><td colspan="2"></td></tr></tfoot></table></div>
      <p class="tenue" style="font-size:12px">Tipos: MO mano de obra · MAT material · EQ equipo · SC subcontrato/servicio · %MO herramientas (% de la mano de obra) · SP subpartida (no se controla). Mano de obra: cantidad (hh) = cuadrilla × ${jor} h / rendimiento.</p>
    </div></div>`;
  enlazarExport(box, { uno: () => rACU([item]) });
  const tb = box.querySelector('#t-acu tbody');
  const refrescarCalc = () => {
    const c2 = S.motor.acu(item);
    tb.querySelectorAll('tr[data-r]').forEach((tr) => {
      const c = c2[Number(tr.dataset.r)];
      if (!c) return;
      tr.querySelector('[data-c=parcial]').textContent = fmt(c.parcial, 4);
      tr.querySelector('[data-c=pct]').textContent = p.pu ? fmt((c.parcial / p.pu) * 100, 1) + ' %' : '';
      const pr = tr.querySelector('[data-c=precio]');
      if (pr) pr.textContent = fmt(c.precio, 4);
    });
    const pu2 = S.motor.puACU(item);
    box.querySelector('#tot-acu').textContent = fmt(pu2, 4);
    box.querySelector('#pu-acu').textContent = fmt(pu2);
    box.querySelector('#pu-acu').classList.toggle('neg', Math.abs(pu2 - p.pu) >= 0.01);
    box.querySelector('#b-usarpu').disabled = Math.abs(pu2 - p.pu) < 0.01;
  };
  tb.addEventListener('change', async (e) => {
    const el = e.target.closest('[data-k]');
    if (!el) return;
    const i = Number(el.closest('tr').dataset.r);
    const f = filas[i];
    const k = el.dataset.k;
    if (k === 'tipo') {
      const nuevo = el.value;
      if (nuevo === '%MO') { f.descripcion = f.insumo || 'HERRAMIENTAS MANUALES'; delete f.insumo; }
      else if (f.tipo === '%MO') { f.insumo = f.descripcion; delete f.descripcion; }
      f.tipo = nuevo;
      const ins = f.insumo && S.motor.insumo(f.insumo);
      if (ins && nuevo !== '%MO' && nuevo !== 'SP') ins.tipo = nuevo;
      cambio(); detalle(box); return;
    }
    if (k === 'insumo') {
      const nom = el.value.trim();
      if (!nom) return;
      f.insumo = nom;
      if (!S.motor.insumo(nom)) {
        const und = tb.querySelector(`tr[data-r="${i}"] [data-k=und]`)?.value || '';
        const pr = num(tb.querySelector(`tr[data-r="${i}"] [data-k=precio]`)?.value);
        o.insumos.push({ codigo: nuevoCodigo(o, f.tipo), descripcion: nom, und, tipo: f.tipo, precio: pr, controlar: true });
        aviso('Insumo nuevo agregado al catálogo: ' + nom);
      }
      cambio(); detalle(box); return;
    }
    if (k === 'und' || k === 'precio') {
      let ins = S.motor.insumo(f.insumo);
      if (!ins) {
        if (!f.insumo) { aviso('Primero escriba el nombre del insumo.', 'alerta'); return; }
        ins = { codigo: nuevoCodigo(o, f.tipo), descripcion: f.insumo, und: '', tipo: f.tipo, precio: 0, controlar: true };
        o.insumos.push(ins);
      }
      if (k === 'precio') {
        const usos = Object.values(o.acu).filter((fs) => fs.some((x) => x.insumo && normal(x.insumo) === normal(ins.descripcion))).length;
        ins.precio = num(el.value);
        el.value = fmt(ins.precio, 2);
        if (usos > 1) aviso(`Precio actualizado en ${usos} A.C.U.`);
      } else ins.und = el.value.trim();
      cambio(); refrescarCalc(); return;
    }
    if (k === 'descripcion') f.descripcion = el.value.trim();
    if (k === 'cantidad') { f.cantidad = num(el.value); el.value = cant(f.cantidad, 6); }
    if (k === 'cuadrilla') {
      f.cuadrilla = el.value.trim() === '' ? null : num(el.value);
      const rd = Number(p.rendimiento) || 0;
      if (f.cuadrilla && rd > 0) {
        f.cantidad = (f.cuadrilla * jor) / rd;
        tb.querySelector(`tr[data-r="${i}"] [data-k=cantidad]`).value = cant(f.cantidad, 6);
      }
    }
    cambio(); refrescarCalc();
  });
  tb.addEventListener('click', async (e) => {
    const b = e.target.closest('[data-del]');
    if (!b) return;
    filas.splice(Number(b.dataset.del), 1);
    cambio(); detalle(box);
  });
  box.querySelector('#rend').onchange = (e) => {
    const rd = e.target.value.trim() === '' ? null : num(e.target.value);
    p.rendimiento = rd;
    let n = 0;
    if (rd > 0) filas.forEach((f) => { if (f.cuadrilla && (f.tipo === 'MO' || f.tipo === 'EQ')) { f.cantidad = (f.cuadrilla * jor) / rd; n++; } });
    cambio();
    if (n) aviso(`${n} cantidad(es) recalculada(s) con la cuadrilla.`);
    detalle(box);
  };
  box.querySelector('#b-add').onclick = () => { filas.push({ tipo: 'MAT', insumo: '', cantidad: 0 }); cambio(); detalle(box); setTimeout(() => { const ins = box.querySelectorAll('[data-k=insumo]'); ins[ins.length - 1]?.focus(); }, 20); };
  box.querySelector('#b-cuad').onclick = () => {
    const rd = Number(p.rendimiento) || 0;
    if (!rd) { aviso('Indique el rendimiento de la partida.', 'alerta'); return; }
    let n = 0;
    filas.forEach((f) => { if ((f.tipo === 'MO' || f.tipo === 'EQ') && f.cuadrilla) { f.cantidad = (f.cuadrilla * jor) / rd; n++; } });
    if (!n) { aviso('Ningún insumo de mano de obra o equipo tiene cuadrilla.', 'alerta'); return; }
    cambio(); detalle(box); aviso(`${n} cantidad(es) recalculada(s).`);
  };
  box.querySelector('#b-usarpu').onclick = async () => {
    const pu2 = S.motor.puACU(item);
    if (!(await confirmar(`¿Cambiar el P.U. de la partida de S/ ${fmt(p.pu)} a S/ ${fmt(pu2)}? El parcial pasa a S/ ${fmt(r2(pu2 * p.metrado))}.`))) return;
    p.pu = pu2; cambio(); refrescar();
  };
  box.querySelector('#b-calc').onclick = () => ir('calculadora', { item });
  box.querySelector('#b-vaciar').onclick = async () => {
    if (!(await confirmar('¿Quitar todos los insumos de este A.C.U.?', { peligro: true }))) return;
    o.acu[item] = []; cambio(); detalle(box);
  };
  box.querySelector('#b-copiar').onclick = async () => {
    const [orig] = await elegirPartidas({ multiple: false, titulo: 'Copiar el A.C.U. de…' });
    if (!orig) return;
    const src = o.acu[orig] || [];
    if (!src.length) { aviso('Esa partida no tiene A.C.U.', 'alerta'); return; }
    if (filas.length && !(await confirmar(`¿Reemplazar los ${filas.length} insumos actuales por los ${src.length} de ${orig}?`))) return;
    o.acu[item] = JSON.parse(JSON.stringify(src));
    cambio(); detalle(box); aviso('A.C.U. copiado de ' + orig);
  };
}
function nuevoCodigo(o, tipo) {
  const p = { MO: 'MO', MAT: 'MT', EQ: 'EQ', SC: 'SC', SP: 'SP' }[tipo] || 'OT';
  let n = 0;
  o.insumos.forEach((i) => { if (i.codigo?.startsWith(p + '-')) n = Math.max(n, parseInt(i.codigo.slice(p.length + 1), 10) || 0); });
  return `${p}-${String(n + 1).padStart(4, '0')}`;
}

export async function importarACU() {
  const f = await elegirArchivo('.xlsx,.xlsm,.csv');
  if (!f) return;
  let res;
  try { res = await trabajando('Leyendo ' + f.name + '…', async () => analizarACU(await leerLibro(f))); } catch (e) { aviso(e.message, 'error', 9000); return; }
  const o = S.obra;
  const items = Object.keys(res.acu);
  const nF = items.reduce((s, k) => s + res.acu[k].length, 0);
  const ej = items.slice(0, 3).map((it) => `<tr class="tit"><td colspan="5">Partida ${esc(it)} ${res.rendimientos[it] ? '· rendimiento ' + res.rendimientos[it] : ''}</td></tr>` + res.acu[it].slice(0, 8).map((x) => `<tr><td>${x.tipo}</td><td>${esc(x.nombre)}</td><td>${esc(x.und)}</td><td class="n">${cant(x.cantidad, 6)}</td><td class="n">${fmt(x.precio)}</td></tr>`).join('')).join('');
  ventana({
    titulo: 'Vista previa de los A.C.U.: ' + f.name, ancho: '900px',
    cuerpo: `<div class="tarjetas"><div class="tarjeta"><div class="et">Hoja</div><div class="val" style="font-size:16px">${esc(res.hoja)}</div></div><div class="tarjeta"><div class="et">Partidas con análisis</div><div class="val">${items.length}</div></div><div class="tarjeta"><div class="et">Filas de insumos</div><div class="val">${nF}</div></div><div class="tarjeta"><div class="et">Con rendimiento</div><div class="val">${Object.keys(res.rendimientos).length}</div></div></div>
      ${res.avisos.length ? `<div class="nota nota-alerta">${res.avisos.map(esc).join('<br>')}</div>` : ''}
      <div class="tabla-cont baja"><table class="t"><thead><tr><th>Tipo</th><th>Insumo</th><th>Und</th><th class="n">Cantidad</th><th class="n">Precio</th></tr></thead><tbody>${ej}</tbody></table></div>
      <div style="margin-top:12px;display:flex;flex-direction:column;gap:6px">
        <label class="chk"><input type="checkbox" id="i-reemp" checked> Reemplazar el A.C.U. de las partidas que ya tienen análisis</label>
        <label class="chk"><input type="checkbox" id="i-precios" checked> Actualizar los precios del catálogo con los del archivo</label>
      </div>`,
    botones: [{ txt: 'Cancelar' }, { txt: 'Importar A.C.U.', clase: 'btn-primario', icono: 'importar', accion: (c, raiz) => {
      const r = aplicarACU(o, res, { reemplazar: raiz.querySelector('#i-reemp').checked, actualizarPrecios: raiz.querySelector('#i-precios').checked });
      cambio({ guardarYa: true });
      aviso(`A.C.U. importados: ${r.partidas} partidas, ${r.filas} filas, ${r.nuevos} insumos nuevos, ${r.precios} precios actualizados.${r.sinPartida.length ? ` ${r.sinPartida.length} análisis sin partida en el presupuesto.` : ''}`, r.sinPartida.length ? 'alerta' : 'ok', 9000);
      refrescar();
    } }],
  });
}
export { rPlantilla };
