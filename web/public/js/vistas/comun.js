// Piezas compartidas por las pantallas.
import { S } from '../store.js';
import { esc, fmt, normal } from '../util.js';
import { icono, ventana, aviso, trabajando } from '../ui.js';
import { exportar } from '../exportar.js';
import { ESTADOS } from '../calc.js';

export function cab(titulo, desc, acciones = '') {
  return `<div class="vista-cab"><div><h1>${esc(titulo)}</h1>${desc ? `<p>${desc}</p>` : ''}</div><div class="acciones">${acciones}</div></div>`;
}
export function btnExport(clave, txt = '') {
  return `<button class="btn" data-exp="pdf" data-rep="${clave}" title="Exportar a PDF">${icono('pdf')}${txt ? esc(txt) + ' ' : ''}PDF</button><button class="btn" data-exp="xlsx" data-rep="${clave}" title="Exportar a Excel con fórmulas">${icono('excel')}${txt ? esc(txt) + ' ' : ''}Excel</button>`;
}
/** Enlaza los botones [data-exp] del contenedor con funciones que devuelven el reporte. */
export function enlazarExport(cont, reportes) {
  cont.querySelectorAll('[data-exp]').forEach((b) => {
    b.addEventListener('click', async () => {
      const fn = reportes[b.dataset.rep];
      if (!fn) return;
      try {
        await trabajando('Generando ' + (b.dataset.exp === 'pdf' ? 'PDF' : 'Excel') + '…', async () => {
          const R = await fn();
          if (!R) return;
          await exportar(R, b.dataset.exp);
        });
      } catch (e) {
        console.error(e);
        aviso('No se pudo generar el archivo: ' + e.message, 'error', 7000);
      }
    });
  });
}
export function chipEstado(e, corto = false) {
  if (!e) return '';
  const est = e.estado || e;
  return `<span class="chip ${est.clase}" title="${esc((e.motivos || []).join('\n'))}">${esc(corto ? est.txt.split(' /')[0] : est.txt)}</span>`;
}
export function leyendaEstados() {
  return `<div class="leyenda">${Object.values(ESTADOS).map((e) => `<span class="chip ${e.clase}">${e.txt}</span>`).join('')}</div>`;
}
export function sinPresupuesto() {
  return `<div class="panel"><div class="vacio"><h3>La obra aún no tiene presupuesto</h3><p>Importe el presupuesto desde Excel (S10, Delphin u otro) o digítelo partida por partida.</p><p><a class="btn btn-primario" href="#/presupuesto">${icono('presupuesto')}Ir al presupuesto</a></p></div></div>`;
}

/** Ventana para buscar y elegir partidas. Devuelve una promesa con los ítems elegidos. */
export function elegirPartidas({ multiple = true, soloP = true, titulo = 'Elegir partida' } = {}) {
  return new Promise((res) => {
    const m = S.motor;
    const lista = S.obra.partidas.filter((p) => !soloP || p.tipo === 'P');
    const sel = new Set();
    let resultado = [];
    const pinta = (raiz, q = '') => {
      const nq = normal(q);
      const f = lista.filter((p) => !nq || normal(p.item + ' ' + p.descripcion).includes(nq)).slice(0, 250);
      raiz.querySelector('#lp').innerHTML = f.map((p) => `<tr data-it="${esc(p.item)}" class="${sel.has(p.item) ? 'sel' : ''} ${p.tipo === 'T' ? 'tit' : ''}" style="cursor:pointer"><td class="item">${esc(p.item)}</td><td>${esc(p.descripcion)}</td><td>${esc(p.und || '')}</td><td class="n">${p.tipo === 'P' ? fmt(p.metrado) : ''}</td><td class="n">${p.tipo === 'P' ? fmt(p.pu) : ''}</td><td class="n">${m.acu(p.item).length || ''}</td></tr>`).join('') || '<tr><td colspan="6" class="tenue">Sin resultados</td></tr>';
      raiz.querySelector('#lp-n').textContent = sel.size ? `${sel.size} seleccionada(s)` : '';
    };
    ventana({
      titulo, ancho: '980px',
      cuerpo: `<div class="filtros"><input type="search" id="lp-q" placeholder="Buscar por ítem o descripción…" autofocus style="flex:1"><span id="lp-n" class="badge badge-azul"></span></div>
        <div class="tabla-cont baja"><table class="t"><thead><tr><th>Ítem</th><th>Descripción</th><th>Und</th><th class="n">Metrado E.T.</th><th class="n">P.U.</th><th class="n">Insumos</th></tr></thead><tbody id="lp"></tbody></table></div>
        <p class="tenue" style="font-size:12px;margin:8px 0 0">${multiple ? 'Haga clic en las filas para marcarlas (puede elegir varias).' : 'Haga clic en una fila para elegirla.'}</p>`,
      botones: multiple ? [{ txt: 'Cancelar' }, { txt: 'Agregar seleccionadas', clase: 'btn-primario', accion: () => { resultado = [...sel]; } }] : [{ txt: 'Cancelar' }],
      alAbrir: (raiz, cerrar) => {
        pinta(raiz);
        raiz.querySelector('#lp-q').addEventListener('input', (e) => pinta(raiz, e.target.value));
        raiz.querySelector('#lp').addEventListener('click', (e) => {
          const tr = e.target.closest('tr[data-it]');
          if (!tr) return;
          const it = tr.dataset.it;
          if (!multiple) { resultado = [it]; cerrar(); return; }
          if (sel.has(it)) sel.delete(it); else sel.add(it);
          pinta(raiz, raiz.querySelector('#lp-q').value);
        });
      },
      alCerrar: () => res(resultado),
    });
  });
}

export function opcionesPeriodo(sel) {
  return S.motor.periodos().map((p, k) => `<option value="${k}" ${String(k) === String(sel) ? 'selected' : ''}>N° ${String(p.numero).padStart(2, '0')} · ${p.del.split('-').reverse().join('/')} – ${p.al.split('-').reverse().join('/')} ${p.estado === 'CERRADO' ? '(cerrada)' : ''}</option>`).join('');
}
export function mesesDeObra() {
  const o = S.obra;
  const set = new Set();
  Object.keys(o.metrados || {}).forEach((f) => set.add(f.slice(0, 7)));
  (o.movimientos || []).forEach((m) => m.fecha && set.add(m.fecha.slice(0, 7)));
  (o.partes || []).forEach((p) => p.fecha && set.add(p.fecha.slice(0, 7)));
  (o.periodos || []).forEach((p) => set.add(p.al.slice(0, 7)));
  if (o.datos?.fecha_inicio) set.add(o.datos.fecha_inicio.slice(0, 7));
  return [...set].sort();
}
export function listaPartidasDatalist(id = 'dl-partidas') {
  return `<datalist id="${id}">${S.obra.partidas.filter((p) => p.tipo === 'P').map((p) => `<option value="${esc(p.item)}">${esc(p.descripcion.slice(0, 80))}</option>`).join('')}</datalist>`;
}
export function datalistInsumos(id = 'dl-insumos', tipos = null) {
  return `<datalist id="${id}">${S.obra.insumos.filter((i) => !tipos || tipos.includes(i.tipo)).map((i) => `<option value="${esc(i.descripcion)}">${esc(i.und)} · S/ ${fmt(i.precio)}</option>`).join('')}</datalist>`;
}
