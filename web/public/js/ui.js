// Componentes de interfaz: avisos, ventanas, formularios, descarga de archivos.
import { esc } from './util.js';

export const $ = (s, r = document) => r.querySelector(s);
export const $$ = (s, r = document) => [...r.querySelectorAll(s)];

// ------------------------------------------------------------------ íconos (SVG en línea, trazo)
const P = {
  panel: 'M3 3h7v9H3zM14 3h7v5h-7zM14 12h7v9h-7zM3 16h7v5H3z',
  obra: 'M3 21h18M5 21V8l7-5 7 5v13M9 21v-6h6v6',
  presupuesto: 'M4 4h16v16H4zM4 9h16M9 9v11',
  acu: 'M9 3h6l1 4H8zM6 7h12v14H6zM9 12h6M9 16h4',
  insumos: 'M21 8l-9-5-9 5 9 5 9-5zM3 8v8l9 5 9-5V8M12 13v8',
  calc: 'M5 3h14v18H5zM8 7h8M8 11h2M12 11h2M16 11v6M8 15h2M12 15h2',
  metrados: 'M3 17l6-6 4 4 8-8M14 7h7v7',
  almacen: 'M3 9l9-6 9 6v12H3zM8 21v-8h8v8',
  equipos: 'M3 17h4l2-5h6l2 5h4M7 17a2 2 0 1 0 4 0M13 17a2 2 0 1 0 4 0M9 12V7h4l2 5',
  valorizacion: 'M12 2v20M17 6H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6',
  cronograma: 'M3 5h18v16H3zM3 10h18M8 3v4M16 3v4M7 14h4M7 17h7',
  reportes: 'M6 2h9l5 5v15H6zM14 2v6h6M9 13h6M9 17h6',
  respaldo: 'M12 3v12M7 10l5 5 5-5M4 21h16',
  mas: 'M12 5v14M5 12h14',
  editar: 'M4 20h4L19 9l-4-4L4 16zM14 6l4 4',
  borrar: 'M4 7h16M9 7V4h6v3M6 7l1 14h10l1-14',
  pdf: 'M6 2h9l5 5v15H6zM14 2v6h6M8 16h2M8 13h8',
  excel: 'M4 4h16v16H4zM9 8l6 8M15 8l-6 8',
  importar: 'M12 21V9M7 14l5-5 5 5M4 3h16',
  buscar: 'M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM21 21l-5-5',
  copiar: 'M8 8h12v12H8zM4 16V4h12',
  candado: 'M6 11h12v10H6zM8 11V7a4 4 0 0 1 8 0v4',
  ok: 'M5 12l5 5L20 7',
  alerta: 'M12 3l10 18H2zM12 10v5M12 18v.01',
  menu: 'M3 6h18M3 12h18M3 18h18',
  cerrar: 'M6 6l12 12M18 6L6 18',
  flecha: 'M9 6l6 6-6 6',
  guardar: 'M5 3h12l4 4v14H5zM8 3v6h8M8 21v-7h8v7',
  ojo: 'M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12zM12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6z',
  usuarios: 'M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM2 21v-1a6 6 0 0 1 12 0v1M16 3.5a4 4 0 0 1 0 7.5M22 21v-1a6 6 0 0 0-4-5.6',
  cuenta: 'M12 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM4 21v-1a7 7 0 0 1 16 0v1',
  salir: 'M15 4h4v16h-4M10 8l-4 4 4 4M6 12h10',
  nube: 'M7 18h10a4 4 0 0 0 .5-8A6 6 0 0 0 6 9.5 4.3 4.3 0 0 0 7 18z',
  llave: 'M15 7a4 4 0 1 1-3.9 5H8v3H5v-3H3V9h8.1A4 4 0 0 1 15 7z',
};
export function icono(n, cls = '') {
  return `<svg class="ico ${cls}" viewBox="0 0 24 24" aria-hidden="true"><path d="${P[n] || ''}"/></svg>`;
}

// ------------------------------------------------------------------ avisos
export function aviso(msg, tipo = 'ok', ms = 3500) {
  let c = $('#avisos');
  if (!c) { c = document.createElement('div'); c.id = 'avisos'; document.body.appendChild(c); }
  const d = document.createElement('div');
  d.className = 'aviso aviso-' + tipo;
  d.innerHTML = `${icono(tipo === 'ok' ? 'ok' : 'alerta')}<span>${esc(msg)}</span>`;
  c.appendChild(d);
  setTimeout(() => { d.classList.add('sale'); setTimeout(() => d.remove(), 300); }, ms);
}

// ------------------------------------------------------------------ ventana modal
let pila = [];
/**
 * Abre una ventana. `cuerpo` es HTML. `botones`: [{txt, clase, accion(cerrar, raiz) → false para no cerrar}]
 * Devuelve {raiz, cerrar}.
 */
export function ventana({ titulo, cuerpo, botones = [{ txt: 'Cerrar', clase: '' }], ancho = '640px', alAbrir, alCerrar }) {
  const fondo = document.createElement('div');
  fondo.className = 'modal-fondo';
  fondo.innerHTML = `<div class="modal" role="dialog" aria-modal="true" style="max-width:${ancho}">
    <div class="modal-cab"><h3>${esc(titulo)}</h3><button class="btn-icono" data-x title="Cerrar">${icono('cerrar')}</button></div>
    <div class="modal-cuerpo">${cuerpo}</div>
    <div class="modal-pie">${botones.map((b, i) => `<button class="btn ${b.clase || ''}" data-b="${i}">${b.icono ? icono(b.icono) : ''}${esc(b.txt)}</button>`).join('')}</div></div>`;
  document.body.appendChild(fondo);
  const raiz = fondo.querySelector('.modal');
  const cerrar = () => {
    fondo.remove();
    pila = pila.filter((x) => x !== fondo);
    alCerrar && alCerrar();
  };
  pila.push(fondo);
  fondo.addEventListener('mousedown', (e) => { if (e.target === fondo) fondo._down = true; });
  fondo.addEventListener('mouseup', (e) => { if (e.target === fondo && fondo._down) cerrar(); fondo._down = false; });
  raiz.querySelector('[data-x]').onclick = cerrar;
  raiz.querySelectorAll('[data-b]').forEach((b) => {
    b.onclick = async () => {
      const def = botones[Number(b.dataset.b)];
      if (!def.accion) return cerrar();
      b.disabled = true;
      try {
        const r = await def.accion(cerrar, raiz);
        if (r !== false) cerrar();
      } catch (err) {
        console.error(err);
        aviso(err.message || String(err), 'error', 6000);
      } finally {
        b.disabled = false;
      }
    };
  });
  setTimeout(() => {
    const f = raiz.querySelector('[autofocus], .modal-cuerpo input:not([type=hidden]):not([readonly]), .modal-cuerpo select, .modal-cuerpo textarea');
    f && f.focus();
    alAbrir && alAbrir(raiz, cerrar);
  }, 20);
  return { raiz, cerrar };
}
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && pila.length) pila[pila.length - 1].querySelector('[data-x]').click();
});

export function confirmar(msg, { titulo = 'Confirmar', si = 'Aceptar', peligro = false } = {}) {
  return new Promise((res) => {
    let r = false;
    ventana({
      titulo, ancho: '460px',
      cuerpo: `<p>${msg}</p>`,
      botones: [{ txt: 'Cancelar' }, { txt: si, clase: peligro ? 'btn-peligro' : 'btn-primario', accion: () => { r = true; } }],
      alCerrar: () => res(r),
    });
  });
}

// ------------------------------------------------------------------ formularios
/**
 * campos: [{id, et, tipo: text|number|date|select|textarea|check, ops:[[v,txt]], v, req, paso, ayuda, col (1|2|3)}]
 */
export function htmlCampos(campos, val = {}) {
  return `<div class="form-grid">${campos.map((c) => {
    if (c.tipo === 'seccion') return `<div class="form-sec">${esc(c.et)}</div>`;
    const v = val[c.id] ?? c.v ?? '';
    const ancho = c.col === 3 ? 'c3' : c.col === 2 ? 'c2' : '';
    let inp;
    if (c.tipo === 'select') inp = `<select name="${c.id}">${c.ops.map((o) => { const [k, t] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(k)}" ${String(k) === String(v) ? 'selected' : ''}>${esc(t)}</option>`; }).join('')}</select>`;
    else if (c.tipo === 'textarea') inp = `<textarea name="${c.id}" rows="${c.filas || 3}">${esc(v)}</textarea>`;
    else if (c.tipo === 'check') inp = `<label class="chk"><input type="checkbox" name="${c.id}" ${v ? 'checked' : ''}> ${esc(c.txt || '')}</label>`;
    else inp = `<input name="${c.id}" type="${c.tipo === 'number' ? 'text' : c.tipo || 'text'}" ${c.tipo === 'number' ? 'inputmode="decimal" data-num' : ''} value="${esc(v)}" ${c.req ? 'required' : ''} ${c.lista ? `list="${c.lista}"` : ''} ${c.ro ? 'readonly' : ''} placeholder="${esc(c.ph || '')}">`;
    return `<label class="campo ${ancho}"><span>${esc(c.et)}${c.req ? ' *' : ''}</span>${inp}${c.ayuda ? `<small>${esc(c.ayuda)}</small>` : ''}</label>`;
  }).join('')}</div>`;
}
export function leerCampos(raiz, campos) {
  const o = {};
  for (const c of campos) {
    if (c.tipo === 'seccion') continue;
    const el = raiz.querySelector(`[name="${c.id}"]`);
    if (!el) continue;
    if (c.tipo === 'check') o[c.id] = el.checked;
    else if (c.tipo === 'number') o[c.id] = el.value.trim() === '' ? null : Number(String(el.value).replace(/,/g, ''));
    else o[c.id] = el.value.trim();
    if (c.req && (o[c.id] === '' || o[c.id] === null)) throw new Error(`Complete el campo «${c.et}».`);
    if (c.tipo === 'number' && o[c.id] !== null && !isFinite(o[c.id])) throw new Error(`«${c.et}» debe ser un número.`);
  }
  return o;
}
export function formulario({ titulo, campos, val = {}, ancho = '720px', guardar = 'Guardar', alGuardar, extra = '' }) {
  return ventana({
    titulo, ancho,
    cuerpo: htmlCampos(campos, val) + extra,
    botones: [{ txt: 'Cancelar' }, { txt: guardar, clase: 'btn-primario', icono: 'guardar', accion: async (cerrar, raiz) => alGuardar(leerCampos(raiz, campos), raiz) }],
  });
}

// ------------------------------------------------------------------ archivos
export function descargar(nombre, datos, tipo = 'application/octet-stream') {
  const blob = datos instanceof Blob ? datos : new Blob([datos], { type: tipo });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1500);
}
export function elegirArchivo(accept = '*') {
  return new Promise((res) => {
    const i = document.createElement('input');
    i.type = 'file';
    i.accept = accept;
    i.style.display = 'none';
    document.body.appendChild(i);
    i.onchange = () => { res(i.files[0] || null); i.remove(); };
    i.click();
  });
}
export function nombreArchivo(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w.-]+/g, '_').replace(/_+/g, '_').slice(0, 90);
}

/** Indicador de trabajo para procesos largos. */
export async function trabajando(msg, fn) {
  const d = document.createElement('div');
  d.className = 'cargando';
  d.innerHTML = `<div class="spinner"></div><span>${esc(msg)}</span>`;
  document.body.appendChild(d);
  await new Promise((r) => setTimeout(r, 30));
  try { return await fn(); } finally { d.remove(); }
}

/** Paginador simple para tablas largas. */
export function paginar(lista, pag, tam) {
  const n = Math.max(1, Math.ceil(lista.length / tam));
  const p = Math.min(Math.max(1, pag), n);
  return { pagina: p, paginas: n, filas: lista.slice((p - 1) * tam, p * tam) };
}
export function htmlPaginador(p, n, total) {
  if (n <= 1) return `<div class="pag"><span>${total} registro(s)</span></div>`;
  return `<div class="pag"><button class="btn btn-sm" data-pag="${p - 1}" ${p <= 1 ? 'disabled' : ''}>‹ Anterior</button><span>Página ${p} de ${n} · ${total} registros</span><button class="btn btn-sm" data-pag="${p + 1}" ${p >= n ? 'disabled' : ''}>Siguiente ›</button></div>`;
}
