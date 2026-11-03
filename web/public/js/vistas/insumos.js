// Catálogo de insumos (precios que usan todos los A.C.U.) y relación total de insumos del presupuesto.
import { S, cambio } from '../store.js';
import { esc, fmt, cant, num, normal } from '../util.js';
import { icono, aviso, confirmar, elegirArchivo, trabajando, paginar, htmlPaginador } from '../ui.js';
import { leerLibro, importarPrecios } from '../importar.js';
import { rInsumos, rExplosion, rPlantilla } from '../reportes.js';
import { TIPOS_INSUMO } from '../calc.js';
import { cab, btnExport, enlazarExport } from './comun.js';
import { refrescar } from '../app.js';

let tab = 'catalogo', q = '', tipo = '', pag = 1;

export function render(cont) {
  const o = S.obra, m = S.motor;
  cont.innerHTML = `
    ${cab('Insumos y precios', 'El precio de cada insumo es único para la obra: si lo cambia aquí, cambian todos los A.C.U. que lo usan (el P.U. del presupuesto no cambia hasta que usted lo actualice).',
      `<button class="btn" id="b-imp">${icono('importar')}Importar precios (Excel)</button><button class="btn" id="b-plant">${icono('excel')}Plantilla</button>`)}
    <div class="pestanas"><button data-tab="catalogo" class="${tab === 'catalogo' ? 'activo' : ''}">Catálogo (${o.insumos.length})</button><button data-tab="explosion" class="${tab === 'explosion' ? 'activo' : ''}">Relación de insumos del presupuesto</button></div>
    <div id="cuerpo"></div>`;
  cont.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { tab = b.dataset.tab; pag = 1; refrescar(); }));
  cont.querySelector('#b-plant').onclick = async () => { const { aExcel } = await import('../exportar.js'); aExcel(rPlantilla('insumos')); };
  cont.querySelector('#b-imp').onclick = async () => {
    const f = await elegirArchivo('.xlsx,.xlsm,.csv');
    if (!f) return;
    try {
      const hojas = await trabajando('Leyendo…', () => leerLibro(f));
      const r = importarPrecios(o, hojas[0].filas);
      cambio({ guardarYa: true });
      aviso(`${r.actualizados} precio(s) actualizado(s), ${r.nuevos} insumo(s) nuevo(s).`);
      refrescar();
    } catch (e) { aviso(e.message, 'error', 8000); }
  };
  const cuerpo = cont.querySelector('#cuerpo');
  if (tab === 'catalogo') catalogo(cuerpo, o, m); else explosion(cuerpo, m);
}

function usos(o) {
  const u = new Map();
  for (const filas of Object.values(o.acu)) for (const f of filas) if (f.insumo) { const k = normal(f.insumo); u.set(k, (u.get(k) || 0) + 1); }
  return u;
}
function catalogo(box, o, m) {
  const u = usos(o);
  const nq = normal(q);
  const lista = o.insumos.map((x, i) => ({ x, i })).filter(({ x }) => (!nq || normal(x.codigo + ' ' + x.descripcion).includes(nq)) && (!tipo || x.tipo === tipo));
  const pg = paginar(lista, pag, 150);
  box.innerHTML = `
    <div class="filtros"><input type="search" class="buscar" id="i-q" placeholder="Buscar insumo o código…" value="${esc(q)}">
      <select id="i-tipo"><option value="">Todos los tipos</option>${Object.entries(TIPOS_INSUMO).filter(([k]) => k !== '%MO').map(([k, t]) => `<option value="${k}" ${tipo === k ? 'selected' : ''}>${k} · ${t}</option>`).join('')}</select>
      <span style="flex:1"></span><button class="btn btn-sm" id="b-nuevo">${icono('mas')}Nuevo insumo</button>${btnExport('cat')}</div>
    <div class="tabla-cont"><table class="t" id="t-ins"><thead><tr><th>Código</th><th>Descripción</th><th>Und</th><th>Tipo</th><th class="n">Precio S/</th><th class="c" title="Se controla con ingresos al almacén / O/S para el estado NO VALORIZAR">Controlar</th><th class="n">Usado en</th><th></th></tr></thead><tbody>
    ${pg.filas.map(({ x, i }) => `<tr data-i="${i}"><td><input data-k="codigo" value="${esc(x.codigo)}" style="width:86px" class="mono"></td><td><input data-k="descripcion" value="${esc(x.descripcion)}" style="min-width:340px"></td><td><input data-k="und" value="${esc(x.und)}" style="width:60px"></td>
      <td><select data-k="tipo">${['MO', 'MAT', 'EQ', 'SC', 'SP'].map((t) => `<option ${x.tipo === t ? 'selected' : ''}>${t}</option>`).join('')}</select></td>
      <td><input class="n" data-k="precio" value="${fmt(x.precio, 2)}" style="width:96px"></td><td class="c"><input type="checkbox" data-k="controlar" ${x.controlar === false ? '' : 'checked'} ${x.tipo === 'MO' ? 'disabled' : ''}></td>
      <td class="n tenue">${u.get(normal(x.descripcion)) || 0} A.C.U.</td><td><button class="btn-icono peligro" data-del="${i}">${icono('borrar')}</button></td></tr>`).join('') || '<tr><td colspan="8" class="vacio">Sin insumos. Se crean al importar o armar los A.C.U.</td></tr>'}
    </tbody></table></div>${htmlPaginador(pg.pagina, pg.paginas, lista.length)}`;
  enlazarExport(box, { cat: () => rInsumos() });
  let t;
  box.querySelector('#i-q').oninput = (e) => { clearTimeout(t); t = setTimeout(() => { q = e.target.value; pag = 1; refrescar().then(() => { const el = document.querySelector('#i-q'); el.focus(); el.setSelectionRange(el.value.length, el.value.length); }); }, 250); };
  box.querySelector('#i-tipo').onchange = (e) => { tipo = e.target.value; pag = 1; refrescar(); };
  box.querySelectorAll('[data-pag]').forEach((b) => (b.onclick = () => { pag = Number(b.dataset.pag); refrescar(); }));
  box.querySelector('#b-nuevo').onclick = () => { o.insumos.unshift({ codigo: 'NV-' + String(o.insumos.length + 1).padStart(4, '0'), descripcion: 'NUEVO INSUMO ' + (o.insumos.length + 1), und: 'und', tipo: 'MAT', precio: 0, controlar: true }); q = ''; tipo = ''; pag = 1; cambio(); refrescar(); };
  box.querySelector('#t-ins tbody').addEventListener('change', (e) => {
    const el = e.target.closest('[data-k]');
    if (!el) return;
    const x = o.insumos[Number(el.closest('tr').dataset.i)];
    const k = el.dataset.k;
    if (k === 'descripcion') {
      const nuevo = el.value.trim();
      if (!nuevo) { el.value = x.descripcion; return; }
      if (normal(nuevo) !== normal(x.descripcion) && m.insumo(nuevo)) { aviso('Ya existe un insumo con ese nombre.', 'error'); el.value = x.descripcion; return; }
      const viejo = normal(x.descripcion);
      for (const filas of Object.values(o.acu)) for (const f of filas) if (f.insumo && normal(f.insumo) === viejo) f.insumo = nuevo;
      for (const a of o.almacen) if (a.insumo_acu && normal(a.insumo_acu) === viejo) a.insumo_acu = nuevo;
      for (const eq of o.equipos) if (eq.insumo_acu && normal(eq.insumo_acu) === viejo) eq.insumo_acu = nuevo;
      x.descripcion = nuevo;
    } else if (k === 'precio') { x.precio = num(el.value); el.value = fmt(x.precio, 2); }
    else if (k === 'controlar') x.controlar = el.checked;
    else x[k] = el.value.trim();
    cambio();
  });
  box.querySelectorAll('[data-del]').forEach((b) => (b.onclick = async () => {
    const x = o.insumos[Number(b.dataset.del)];
    const n = u.get(normal(x.descripcion)) || 0;
    if (!(await confirmar(`¿Eliminar el insumo <b>${esc(x.descripcion)}</b>?${n ? `<br>Se usa en ${n} A.C.U.: también se quitará de esos análisis.` : ''}`, { peligro: true }))) return;
    const k = normal(x.descripcion);
    for (const it of Object.keys(o.acu)) o.acu[it] = o.acu[it].filter((f) => !f.insumo || normal(f.insumo) !== k);
    o.insumos.splice(Number(b.dataset.del), 1);
    cambio(); refrescar();
  }));
}

function explosion(box, m) {
  const ex = m.explosion();
  const tot = ex.reduce((s, x) => s + x.monto, 0);
  const porTipo = {};
  ex.forEach((x) => (porTipo[x.tipo] = (porTipo[x.tipo] || 0) + x.monto));
  box.innerHTML = `
    <div class="tarjetas">${Object.entries(porTipo).map(([t, v]) => `<div class="tarjeta"><div class="et">${esc(TIPOS_INSUMO[t] || t)}</div><div class="val">S/ ${fmt(v)}</div><div class="sub">${fmt((v / (tot || 1)) * 100, 1)} % del total</div></div>`).join('')}<div class="tarjeta"><div class="et">Total (sin %MO)</div><div class="val">S/ ${fmt(tot)}</div></div></div>
    <div class="filtros"><span class="tenue">Cantidad total de cada insumo = Σ metrado del expediente × cantidad del A.C.U. Sirve para el requerimiento general y el cuadro de necesidades.</span><span style="flex:1"></span>${btnExport('exp')}</div>
    <div class="tabla-cont"><table class="t"><thead><tr><th>Código</th><th>Insumo</th><th>Und</th><th>Tipo</th><th class="n">Cantidad</th><th class="n">Precio S/</th><th class="n">Parcial S/</th><th class="n">Partidas</th></tr></thead><tbody>
    ${ex.map((x) => `<tr><td class="mono tenue">${esc(x.codigo)}</td><td>${esc(x.insumo)}</td><td>${esc(x.und)}</td><td>${x.tipo}</td><td class="n">${cant(x.cantidad, 4)}</td><td class="n">${fmt(x.precio)}</td><td class="n">${fmt(x.monto)}</td><td class="n tenue">${x.detalle.length}</td></tr>`).join('')}
    </tbody><tfoot><tr class="tot"><td colspan="6">TOTAL</td><td class="n">${fmt(tot)}</td><td></td></tr></tfoot></table></div>`;
  enlazarExport(box, { exp: () => rExplosion() });
}
