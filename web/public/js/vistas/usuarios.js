// Usuarios y accesos (solo administradores, modo nube): crear cuentas, rol, activo, contraseña y obras de cada cliente.
import { S } from '../store.js';
import { esc, fFecha } from '../util.js';
import { icono, ventana, confirmar, aviso, htmlCampos, leerCampos } from '../ui.js';
import * as nube from '../nube.js';
import { cab } from './comun.js';
import { refrescar } from '../app.js';

let tab = 'usuarios';
let cache = null;

async function cargar() {
  const [usuarios, accesos] = await Promise.all([nube.listarUsuarios(), nube.listarAccesos()]);
  cache = { usuarios, accesos, obras: S.obras };
  return cache;
}
function claveAleatoria() {
  const a = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const v = new Uint32Array(10);
  crypto.getRandomValues(v);
  return 'Sgo-' + [...v].map((x) => a[x % a.length]).join('');
}

export async function render(cont) {
  cont.innerHTML = cab('Usuarios y accesos', 'Cargando…');
  let d;
  try { d = await cargar(); } catch (e) {
    cont.innerHTML = cab('Usuarios y accesos', '') + `<div class="nota nota-error">${esc(e.message)}</div>`;
    return;
  }
  const admins = d.usuarios.filter((u) => u.rol === 'admin');
  const clientes = d.usuarios.filter((u) => u.rol === 'cliente');
  const nObras = (uid) => d.accesos.filter((a) => a.usuario_id === uid).length;
  cont.innerHTML = `
    ${cab('Usuarios y accesos', `${admins.length} administrador(es) y ${clientes.length} cliente(s). El <b>administrador</b> registra y modifica todo; el <b>cliente</b> solo consulta las obras que usted le asigne y descarga sus reportes.`,
      `<button class="btn btn-primario" id="b-nuevo">${icono('mas')}Nuevo usuario</button>`)}
    <div class="pestanas"><button data-tab="usuarios" class="${tab === 'usuarios' ? 'activo' : ''}">Usuarios</button><button data-tab="accesos" class="${tab === 'accesos' ? 'activo' : ''}">Obras de cada cliente</button></div>
    <div id="cuerpo"></div>`;
  cont.querySelectorAll('[data-tab]').forEach((b) => (b.onclick = () => { tab = b.dataset.tab; refrescar(); }));
  cont.querySelector('#b-nuevo').onclick = () => nuevoUsuario();
  const box = cont.querySelector('#cuerpo');

  if (tab === 'usuarios') {
    box.innerHTML = `<div class="tabla-cont libre"><table class="t" id="t-us"><thead><tr><th>Nombre</th><th>Correo</th><th>Rol</th><th class="c">Activo</th><th class="n">Obras asignadas</th><th>Alta</th><th></th></tr></thead><tbody>
      ${d.usuarios.map((u) => { const yo = u.id === S.usuario.id; return `<tr data-id="${esc(u.id)}">
        <td><input data-k="nombre" value="${esc(u.nombre)}" placeholder="Nombre y apellidos" style="min-width:200px"> ${yo ? '<span class="badge badge-azul">usted</span>' : ''}</td>
        <td class="mono" style="font-size:12px">${esc(u.email)}${u.debe_cambiar_clave ? ' <span class="badge badge-alerta" title="Debe cambiar la contraseña temporal al ingresar">clave temporal</span>' : ''}</td>
        <td><select data-k="rol" ${yo ? 'disabled title="No puede cambiar su propio rol"' : ''}><option value="admin" ${u.rol === 'admin' ? 'selected' : ''}>Administrador</option><option value="cliente" ${u.rol === 'cliente' ? 'selected' : ''}>Cliente (solo lectura)</option></select></td>
        <td class="c"><input type="checkbox" data-k="activo" ${u.activo ? 'checked' : ''} ${yo ? 'disabled' : ''}></td>
        <td class="n">${u.rol === 'admin' ? '<span class="tenue">todas</span>' : nObras(u.id)}</td>
        <td>${fFecha(String(u.creado).slice(0, 10))}</td>
        <td class="acc"><button class="btn btn-sm" data-acc="clave" title="Restablecer contraseña">${icono('llave')}Contraseña</button> ${yo ? '' : `<button class="btn btn-sm btn-peligro" data-acc="quitar" title="Quitar el acceso a SGO">${icono('borrar')}Quitar</button>`}</td></tr>`; }).join('')}
      </tbody></table></div>
      <p class="tenue" style="font-size:12px">Las cuentas de correo se comparten con otra aplicación del mismo proyecto de Supabase: «Quitar» solo retira el acceso a SGO (no borra la cuenta) y la contraseña es la misma en ambas.</p>`;
    const tb = box.querySelector('#t-us tbody');
    tb.addEventListener('change', async (e) => {
      const el = e.target.closest('[data-k]');
      if (!el) return;
      const id = el.closest('tr').dataset.id;
      const k = el.dataset.k;
      const valor = k === 'activo' ? el.checked : el.value.trim();
      try {
        await nube.actualizarUsuario(id, { [k]: valor });
        aviso('Usuario actualizado.');
        if (k !== 'nombre') refrescar();
      } catch (err) { aviso(err.message, 'error', 7000); refrescar(); }
    });
    tb.addEventListener('click', async (e) => {
      const b = e.target.closest('[data-acc]');
      if (!b) return;
      const u = d.usuarios.find((x) => x.id === b.closest('tr').dataset.id);
      if (b.dataset.acc === 'clave') restablecerClave(u);
      if (b.dataset.acc === 'quitar') {
        if (!(await confirmar(`¿Quitar el acceso a SGO de <b>${esc(u.nombre || u.email)}</b>? Ya no podrá ingresar al sistema ni ver sus obras. (La cuenta de correo no se borra.)`, { peligro: true, si: 'Quitar acceso' }))) return;
        try { await nube.funcionUsuarios({ accion: 'quitar', id: u.id }); aviso('Acceso retirado.'); refrescar(); } catch (err) { aviso(err.message, 'error', 8000); }
      }
    });
    return;
  }

  // ---- matriz clientes × obras
  if (!clientes.length) { box.innerHTML = `<div class="panel"><div class="vacio"><h3>No hay clientes</h3><p>Cree un usuario con rol «Cliente» y luego marque aquí las obras que puede consultar.</p></div></div>`; return; }
  if (!d.obras.length) { box.innerHTML = `<div class="panel"><div class="vacio"><h3>No hay obras</h3><p>Cree o importe una obra primero.</p></div></div>`; return; }
  const tiene = new Set(d.accesos.map((a) => a.obra_id + '|' + a.usuario_id));
  box.innerHTML = `<p class="tenue" style="margin-top:0">Marque las obras que puede consultar cada cliente. Los administradores ven todas.</p>
    <div class="tabla-cont libre"><table class="t matriz" id="t-acc"><thead><tr><th>Cliente</th>${d.obras.map((o) => `<th class="c" title="${esc(o.nombre)}">${esc(o.nombre.length > 28 ? o.nombre.slice(0, 27) + '…' : o.nombre)}${o.cui ? `<div class="tenue" style="font-weight:400;font-size:10px">CUI ${esc(o.cui)}</div>` : ''}</th>`).join('')}</tr></thead><tbody>
    ${clientes.map((u) => `<tr data-id="${esc(u.id)}"><td><b>${esc(u.nombre || u.email)}</b><div class="tenue mono" style="font-size:11px">${esc(u.email)}</div>${u.activo ? '' : '<span class="badge badge-error">desactivado</span>'}</td>
      ${d.obras.map((o) => `<td class="c"><input type="checkbox" data-obra="${esc(o.id)}" ${tiene.has(o.id + '|' + u.id) ? 'checked' : ''}></td>`).join('')}</tr>`).join('')}
    </tbody></table></div>`;
  box.querySelector('#t-acc tbody').addEventListener('change', async (e) => {
    const el = e.target.closest('[data-obra]');
    if (!el) return;
    const uid = el.closest('tr').dataset.id, oid = el.dataset.obra;
    el.disabled = true;
    try {
      if (el.checked) { await nube.darAcceso(oid, uid); aviso('Acceso concedido.'); }
      else { await nube.quitarAcceso(oid, uid); aviso('Acceso retirado.'); }
    } catch (err) { aviso(err.message, 'error', 7000); el.checked = !el.checked; }
    el.disabled = false;
  });
}

function nuevoUsuario() {
  const campos = [
    { id: 'nombre', et: 'Nombre y apellidos', req: true, col: 2 },
    { id: 'rol', et: 'Rol', tipo: 'select', ops: [['cliente', 'Cliente (solo lectura)'], ['admin', 'Administrador']] },
    { id: 'email', et: 'Correo electrónico', req: true, col: 2 },
    { id: 'clave', et: 'Contraseña temporal', v: claveAleatoria(), ayuda: 'Mínimo 8 caracteres. Se le pedirá cambiarla al ingresar.' },
  ];
  ventana({
    titulo: 'Nuevo usuario', ancho: '720px',
    cuerpo: htmlCampos(campos) + `<label class="chk" style="margin-top:10px"><input type="checkbox" id="nu-rest"> Si el correo ya tiene cuenta en este proyecto, restablecer su contraseña con la temporal</label>
      <p class="tenue" style="font-size:12px">Si el correo ya existe (por ejemplo, en la otra aplicación del proyecto), se le da acceso a SGO y conserva su contraseña actual, salvo que marque la casilla anterior.</p>`,
    botones: [{ txt: 'Cancelar' }, { txt: 'Crear usuario', clase: 'btn-primario', icono: 'guardar', accion: async (cerrar, raiz) => {
      const v = leerCampos(raiz, campos);
      if (v.clave && v.clave.length < 8) throw new Error('La contraseña temporal debe tener al menos 8 caracteres.');
      const r = await nube.funcionUsuarios({ accion: 'crear', email: v.email, nombre: v.nombre, rol: v.rol, clave: v.clave, restablecer: raiz.querySelector('#nu-rest').checked });
      mostrarCredenciales({ ...v, existia: r.existia, temporal: r.temporal });
      refrescar();
    } }],
  });
}
function mostrarCredenciales(v) {
  const url = location.origin + location.pathname;
  const texto = `SGO – Costos y valorizaciones\nIngreso: ${url}\nCorreo: ${v.email}\n${v.temporal ? `Contraseña temporal: ${v.clave}\n(Al ingresar se le pedirá crear su propia contraseña.)` : 'Contraseña: la misma que ya usa con este correo.'}\nRol: ${v.rol === 'admin' ? 'Administrador' : 'Cliente (solo lectura)'}`;
  ventana({
    titulo: v.existia ? 'Acceso concedido a una cuenta existente' : 'Usuario creado', ancho: '560px',
    cuerpo: `<p>${v.existia ? 'El correo ya tenía cuenta en el proyecto: ahora también tiene acceso a SGO.' : 'Envíe estos datos a la persona por un medio seguro (en persona o por mensaje directo).'}</p>
      <textarea id="cred" rows="7" readonly style="width:100%;font-family:var(--mono);font-size:12.5px">${esc(texto)}</textarea>
      ${v.rol === 'cliente' ? '<p class="tenue" style="font-size:12px">Ahora asígnele sus obras en la pestaña «Obras de cada cliente».</p>' : ''}`,
    botones: [{ txt: 'Copiar', clase: 'btn-primario', accion: async (c, raiz) => { try { await navigator.clipboard.writeText(texto); aviso('Copiado.'); } catch { raiz.querySelector('#cred').select(); document.execCommand('copy'); } return false; } }, { txt: 'Listo' }],
  });
}
function restablecerClave(u) {
  const campos = [{ id: 'clave', et: 'Nueva contraseña temporal', v: claveAleatoria(), req: true, ayuda: 'Mínimo 8 caracteres. Se le pedirá cambiarla al ingresar.' }];
  ventana({
    titulo: 'Restablecer contraseña: ' + (u.nombre || u.email), ancho: '520px',
    cuerpo: htmlCampos(campos) + '<p class="tenue" style="font-size:12px">La contraseña es la misma para la otra aplicación que comparte este proyecto de Supabase.</p>',
    botones: [{ txt: 'Cancelar' }, { txt: 'Restablecer', clase: 'btn-primario', accion: async (c, raiz) => {
      const v = leerCampos(raiz, campos);
      if (v.clave.length < 8) throw new Error('Mínimo 8 caracteres.');
      await nube.funcionUsuarios({ accion: 'clave', id: u.id, clave: v.clave });
      mostrarCredenciales({ email: u.email, clave: v.clave, rol: u.rol, temporal: true });
      refrescar();
    } }],
  });
}
