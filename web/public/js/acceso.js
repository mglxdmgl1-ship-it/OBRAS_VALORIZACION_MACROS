// Pantallas de acceso: ingreso, cambio de contraseña obligatorio y cuenta sin acceso a SGO.
import { esc } from './util.js';
import { icono } from './ui.js';
import { ingresar, cambiarMiClave, salir } from './nube.js';

const LOGO = `<svg viewBox="0 0 32 32" class="acceso-logo"><rect width="32" height="32" rx="7" fill="#1f3a5f"/><path d="M7 24h18M9 24V13l7-5 7 5v11M13 24v-6h6v6" stroke="#f4b942" stroke-width="2.2" fill="none" stroke-linejoin="round"/></svg>`;
const raiz = () => document.getElementById('acceso');

function mostrar(html) {
  const r = raiz();
  document.getElementById('app').hidden = true;
  r.hidden = false;
  r.innerHTML = `<div class="acceso-fondo"><div class="acceso-caja">
    <div class="acceso-marca">${LOGO}<div><strong>SGO</strong><span>Costos · Presupuestos · Valorizaciones</span></div></div>
    ${html}
  </div><p class="acceso-pie">Obras por administración directa · Gobierno Regional de Puno</p></div>`;
  return r;
}
export function ocultarAcceso() {
  const r = raiz();
  r.hidden = true;
  r.innerHTML = '';
  document.getElementById('app').hidden = false;
}

/** Formulario de ingreso. Se resuelve con la sesión cuando las credenciales son correctas. */
export function pantallaIngreso(mensaje = '') {
  return new Promise((resolver) => {
    const r = mostrar(`
      <h1>Iniciar sesión</h1>
      <p class="tenue">Ingrese con el correo y la contraseña que le dio el administrador.</p>
      <form id="f-ingreso" autocomplete="on" novalidate>
        <label class="campo"><span>Correo electrónico</span><input name="email" type="email" autocomplete="username" required autofocus></label>
        <label class="campo"><span>Contraseña</span><div class="clave-caja"><input name="clave" type="password" autocomplete="current-password" required><button type="button" class="btn-icono" data-ver title="Mostrar u ocultar la contraseña">${icono('ojo')}</button></div></label>
        <div class="acceso-error" id="ing-error" ${mensaje ? '' : 'hidden'}>${esc(mensaje)}</div>
        <button class="btn btn-primario acceso-btn" type="submit" id="ing-btn">Ingresar</button>
      </form>
      <div class="acceso-roles">
        <div><b>Administrador</b><span>Registra el presupuesto, los A.C.U., los metrados, el almacén y las valorizaciones; crea usuarios y asigna obras.</span></div>
        <div><b>Cliente</b><span>Consulta el avance de sus obras y descarga los reportes en PDF y Excel (solo lectura).</span></div>
      </div>
      <p class="tenue acceso-nota">¿No tiene cuenta u olvidó su contraseña? Pídala al administrador del sistema.</p>`);
    const f = r.querySelector('#f-ingreso');
    const err = r.querySelector('#ing-error');
    const btn = r.querySelector('#ing-btn');
    r.querySelector('[data-ver]').onclick = () => { const i = f.clave; i.type = i.type === 'password' ? 'text' : 'password'; };
    f.addEventListener('submit', async (e) => {
      e.preventDefault();
      const email = f.email.value.trim(), clave = f.clave.value;
      if (!email || !clave) { err.textContent = 'Escriba su correo y su contraseña.'; err.hidden = false; return; }
      btn.disabled = true; btn.textContent = 'Ingresando…'; err.hidden = true;
      try {
        const ses = await ingresar(email, clave);
        resolver(ses);
      } catch (ex) {
        err.textContent = ex.message; err.hidden = false;
        btn.disabled = false; btn.textContent = 'Ingresar';
        f.clave.select();
      }
    });
  });
}

/** Cambio de contraseña (obligatorio después de una contraseña temporal). */
export function pantallaCambioClave({ obligatorio = true, email = '' } = {}) {
  return new Promise((resolver) => {
    const r = mostrar(`
      <h1>${obligatorio ? 'Cree su contraseña' : 'Cambiar contraseña'}</h1>
      <p class="tenue">${obligatorio ? 'Ingresó con una contraseña temporal. Por seguridad, escriba una nueva contraseña personal' : 'Escriba la nueva contraseña'}${email ? ` para <b>${esc(email)}</b>` : ''}.</p>
      <form id="f-clave" novalidate>
        <label class="campo"><span>Nueva contraseña (mínimo 8 caracteres)</span><input name="c1" type="password" autocomplete="new-password" required autofocus></label>
        <label class="campo"><span>Repita la contraseña</span><input name="c2" type="password" autocomplete="new-password" required></label>
        <div class="acceso-error" id="cl-error" hidden></div>
        <button class="btn btn-primario acceso-btn" type="submit" id="cl-btn">Guardar contraseña</button>
        <button class="btn acceso-btn" type="button" id="cl-salir">Salir</button>
      </form>`);
    const f = r.querySelector('#f-clave');
    const err = r.querySelector('#cl-error');
    const btn = r.querySelector('#cl-btn');
    r.querySelector('#cl-salir').onclick = async () => { await salir(); location.reload(); };
    f.addEventListener('submit', async (e) => {
      e.preventDefault();
      const c1 = f.c1.value, c2 = f.c2.value;
      const fallo = c1.length < 8 ? 'La contraseña debe tener al menos 8 caracteres.' : c1 !== c2 ? 'Las dos contraseñas no coinciden.' : '';
      if (fallo) { err.textContent = fallo; err.hidden = false; return; }
      btn.disabled = true; btn.textContent = 'Guardando…'; err.hidden = true;
      try { await cambiarMiClave(c1); resolver(true); }
      catch (ex) { err.textContent = ex.message; err.hidden = false; btn.disabled = false; btn.textContent = 'Guardar contraseña'; }
    });
  });
}

/** La cuenta existe pero no tiene acceso a SGO (o fue desactivada). */
export function pantallaSinAcceso({ email = '', desactivado = false, error = '' } = {}) {
  const r = mostrar(`
    <h1>${error ? 'No se pudo cargar su cuenta' : 'Sin acceso a SGO'}</h1>
    <p>${error ? esc(error) : desactivado ? `La cuenta <b>${esc(email)}</b> está <b>desactivada</b> en SGO.` : `La cuenta <b>${esc(email)}</b> no tiene acceso a SGO.`}</p>
    <p class="tenue">${error ? 'Revise su conexión e intente de nuevo.' : 'Pida al administrador que le dé acceso desde «Usuarios y accesos».'}</p>
    <div class="acciones" style="margin-top:14px">${error ? '<button class="btn btn-primario" id="sa-reint">Reintentar</button>' : ''}<button class="btn" id="sa-salir">Salir</button></div>`);
  r.querySelector('#sa-salir').onclick = async () => { await salir(); location.reload(); };
  r.querySelector('#sa-reint')?.addEventListener('click', () => location.reload());
}

/** Pantalla de carga mientras se verifica la sesión. */
export function pantallaCargando(txt = 'Verificando la sesión…') {
  mostrar(`<div class="acceso-cargando"><div class="spinner"></div><span>${esc(txt)}</span></div>`);
}
