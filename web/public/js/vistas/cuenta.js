// Mi cuenta: datos del usuario de la sesión y cambio de contraseña.
import { S } from '../store.js';
import { esc } from '../util.js';
import { icono, aviso } from '../ui.js';
import { cambiarMiClave } from '../nube.js';
import { cab } from './comun.js';

export function render(cont) {
  const u = S.usuario || {};
  cont.innerHTML = `
    ${cab('Mi cuenta', 'Datos de su usuario en SGO.')}
    <div class="rejilla-2">
      <div class="panel"><div class="panel-cab"><h2>Usuario</h2></div><div class="panel-cuerpo"><div class="kv">
        <div>Nombre</div><div><b>${esc(u.nombre || '—')}</b></div>
        <div>Correo</div><div class="mono">${esc(u.email || '')}</div>
        <div>Rol</div><div>${u.rol === 'admin' ? '<span class="badge badge-azul">Administrador</span> registra y modifica todas las obras, crea usuarios y asigna accesos.' : '<span class="badge badge-alerta">Cliente</span> consulta las obras asignadas y descarga sus reportes (solo lectura).'}</div>
        <div>Obras disponibles</div><div>${S.obras.length}</div>
      </div><p class="tenue" style="font-size:12px;margin-bottom:0">Para cambiar su nombre o su rol, pídalo a un administrador.</p></div></div>
      <div class="panel" data-libre><div class="panel-cab"><h2>Cambiar contraseña</h2></div><div class="panel-cuerpo">
        <form id="f-clave" class="form-grid" style="grid-template-columns:1fr" data-libre>
          <label class="campo"><span>Nueva contraseña (mínimo 8 caracteres)</span><input name="c1" type="password" autocomplete="new-password" data-libre></label>
          <label class="campo"><span>Repita la contraseña</span><input name="c2" type="password" autocomplete="new-password" data-libre></label>
          <div><button class="btn btn-primario" type="submit" data-libre>${icono('llave')}Cambiar contraseña</button></div>
        </form>
        <p class="tenue" style="font-size:12px;margin-bottom:0">La cuenta se comparte con otra aplicación del mismo proyecto: la nueva contraseña sirve para ambas.</p>
      </div></div>
    </div>`;
  const f = cont.querySelector('#f-clave');
  f.addEventListener('submit', async (e) => {
    e.preventDefault();
    const c1 = f.c1.value, c2 = f.c2.value;
    if (c1.length < 8) { aviso('La contraseña debe tener al menos 8 caracteres.', 'error'); return; }
    if (c1 !== c2) { aviso('Las dos contraseñas no coinciden.', 'error'); return; }
    const b = f.querySelector('button');
    b.disabled = true;
    try { await cambiarMiClave(c1); aviso('Contraseña cambiada.'); f.reset(); } catch (err) { aviso(err.message, 'error', 7000); }
    b.disabled = false;
  });
}
