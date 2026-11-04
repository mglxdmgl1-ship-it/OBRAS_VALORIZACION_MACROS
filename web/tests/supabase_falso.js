// Supabase simulado para las pruebas en el navegador (reemplaza vendor/supabase.js).
// Reproduce lo que usa nube.js: auth (sesión, ingreso, cambio de clave), consultas a las tablas sgo_*
// con las mismas reglas de seguridad por fila (RLS), los disparadores de sgo_obras, la regla del último
// administrador, la función sgo-usuarios y el rpc sgo_clave_cambiada. Los datos viven en localStorage.
(function () {
  const KEY = 'sgo-falso-db';
  const ahora = () => new Date().toISOString();
  const semilla = () => ({
    users: [{ id: 'u-admin', email: 'admin@prueba.pe', password: 'Admin12345' }],
    sgo_usuarios: [{ id: 'u-admin', email: 'admin@prueba.pe', nombre: 'ADMINISTRADOR DE PRUEBA', rol: 'admin', activo: true, debe_cambiar_clave: false, creado: ahora() }],
    sgo_obras: [], sgo_accesos: [], sgo_respaldos: [], session: null, seq: 1, llamadas: [],
  });
  const cargar = () => { try { return JSON.parse(localStorage.getItem(KEY)) || semilla(); } catch { return semilla(); } };
  const guardar = (db) => localStorage.setItem(KEY, JSON.stringify(db));
  const espera = (ms = 25) => new Promise((r) => setTimeout(r, ms));
  const clon = (x) => JSON.parse(JSON.stringify(x));
  const err = (message, code) => ({ data: null, error: { message, code } });

  function createClient() {
    const escuchas = [];
    const uid = () => cargar().session?.user?.id || null;
    const perfil = (db) => db.sgo_usuarios.find((u) => u.id === uid());
    const esAdmin = (db) => { const p = perfil(db); return !!p && p.rol === 'admin' && p.activo; };
    const puedeVer = (db, obra) => esAdmin(db) || (db.sgo_accesos.some((a) => a.obra_id === obra && a.usuario_id === uid()) && !!perfil(db)?.activo);
    const RLS = {
      sgo_usuarios: { select: (db, r) => r.id === uid() || esAdmin(db), update: esAdmin, insert: () => false, delete: () => false, cols: ['nombre', 'rol', 'activo', 'debe_cambiar_clave'] },
      sgo_obras: { select: (db, r) => puedeVer(db, r.id), insert: esAdmin, update: esAdmin, delete: esAdmin, cols: ['datos'] },
      sgo_accesos: { select: (db, r) => r.usuario_id === uid() || esAdmin(db), insert: esAdmin, delete: esAdmin, update: () => false },
      sgo_respaldos: { select: esAdmin, insert: () => false, update: () => false, delete: () => false },
    };
    const auth = {
      async getSession() { await espera(); return { data: { session: cargar().session }, error: null }; },
      async signInWithPassword({ email, password }) {
        await espera();
        const db = cargar();
        const u = db.users.find((x) => x.email === String(email).trim().toLowerCase() && x.password === password);
        if (!u) return { data: {}, error: { message: 'Invalid login credentials', status: 400 } };
        db.session = { access_token: 'tok-' + u.id, user: { id: u.id, email: u.email } };
        guardar(db);
        escuchas.forEach((f) => f('SIGNED_IN', db.session));
        return { data: { session: db.session, user: db.session.user }, error: null };
      },
      async signOut() { const db = cargar(); db.session = null; guardar(db); escuchas.forEach((f) => f('SIGNED_OUT', null)); return { error: null }; },
      async updateUser({ password }) {
        await espera();
        const db = cargar();
        if (!db.session) return { data: {}, error: { message: 'Auth session missing!' } };
        if (!password || password.length < 8) return { data: {}, error: { message: 'Password should be at least 8 characters.' } };
        const u = db.users.find((x) => x.id === db.session.user.id);
        if (u.password === password) return { data: {}, error: { message: 'New password should be different from the old password.' } };
        u.password = password;
        guardar(db);
        return { data: { user: db.session.user }, error: null };
      },
      onAuthStateChange(fn) { escuchas.push(fn); return { data: { subscription: { unsubscribe() {} } } }; },
    };

    class Consulta {
      constructor(tabla) { this.t = tabla; this.op = 'select'; this.f = []; this.o = []; this.devolver = false; this.uno = null; }
      select() { if (this.op !== 'select') this.devolver = true; return this; }
      insert(v) { this.op = 'insert'; this.v = v; return this; }
      update(v) { this.op = 'update'; this.v = v; return this; }
      delete() { this.op = 'delete'; return this; }
      eq(c, v) { this.f.push([c, v]); return this; }
      order(c, op) { this.o.push([c, op?.ascending !== false]); return this; }
      maybeSingle() { this.uno = 'maybe'; return this; }
      single() { this.uno = 'single'; return this; }
      then(ok, mal) { return this.ejecutar().then(ok, mal); }
      async ejecutar() {
        await espera();
        const db = cargar();
        db.llamadas.push(this.op + ' ' + this.t);
        if (!db.session) { guardar(db); return err('JWT expired', 'PGRST301'); }
        const pol = RLS[this.t];
        if (!pol) return err(`relation "public.${this.t}" does not exist`, '42P01');
        const filas = db[this.t];
        const coincide = (r) => this.f.every(([c, v]) => r[c] === v);
        let res;
        if (this.op === 'select') {
          res = filas.filter((r) => pol.select(db, r) && coincide(r));
          for (const [c, asc] of [...this.o].reverse()) res.sort((a, b) => ((a[c] > b[c]) - (a[c] < b[c])) * (asc ? 1 : -1));
          res = clon(res);
        } else if (this.op === 'insert') {
          if (!pol.insert(db)) return err('new row violates row-level security policy for table "' + this.t + '"', '42501');
          const nuevos = [].concat(this.v).map(clon);
          for (const n of nuevos) {
            if (this.t === 'sgo_obras') {
              if (!/^[A-Za-z0-9._-]{3,80}$/.test(n.id)) return err('new row violates check constraint', '23514');
              if (filas.some((r) => r.id === n.id)) return err('duplicate key value violates unique constraint "sgo_obras_pkey"', '23505');
              Object.assign(n, { version: 1, creado: ahora(), creado_por: uid(), actualizado: ahora(), actualizado_por: uid(), nombre: n.datos?.datos?.nombre_corto || n.datos?.datos?.nombre || 'SIN NOMBRE', cui: n.datos?.datos?.cui || '' });
              n.datos.id = n.id;
            }
            if (this.t === 'sgo_accesos') {
              if (filas.some((r) => r.obra_id === n.obra_id && r.usuario_id === n.usuario_id)) return err('duplicate key value violates unique constraint', '23505');
              n.creado = ahora();
            }
            filas.push(n);
          }
          res = clon(nuevos);
        } else if (this.op === 'update') {
          const prohib = Object.keys(this.v).filter((k) => !(pol.cols || []).includes(k));
          if (prohib.length) return err('permission denied for table ' + this.t, '42501');
          const objetivo = pol.update(db) ? filas.filter((r) => pol.select(db, r) && coincide(r)) : [];
          for (const r of objetivo) {
            if (this.t === 'sgo_usuarios' && r.rol === 'admin' && r.activo && (this.v.rol && this.v.rol !== 'admin' || this.v.activo === false)) {
              if (!filas.some((x) => x.id !== r.id && x.rol === 'admin' && x.activo)) return err('Debe quedar al menos un administrador activo.', 'P0001');
            }
            if (this.t === 'sgo_obras') {
              const dia = db.sgo_respaldos.some((x) => x.obra_id === r.id && Date.now() - Date.parse(x.creado) < 20 * 3600e3);
              if (!dia) db.sgo_respaldos.push({ id: db.seq++, obra_id: r.id, version: r.version, datos: clon(r.datos), creado: ahora(), creado_por: r.actualizado_por });
              Object.assign(r, clon(this.v), { version: r.version + 1, actualizado: ahora(), actualizado_por: uid() });
              r.datos.id = r.id;
              r.nombre = r.datos?.datos?.nombre_corto || r.datos?.datos?.nombre || 'SIN NOMBRE';
              r.cui = r.datos?.datos?.cui || '';
            } else Object.assign(r, clon(this.v));
          }
          res = clon(objetivo);
        } else if (this.op === 'delete') {
          const objetivo = pol.delete(db) ? filas.filter((r) => pol.select(db, r) && coincide(r)) : [];
          db[this.t] = filas.filter((r) => !objetivo.includes(r));
          if (this.t === 'sgo_obras') { const ids = new Set(objetivo.map((r) => r.id)); db.sgo_accesos = db.sgo_accesos.filter((a) => !ids.has(a.obra_id)); db.sgo_respaldos = db.sgo_respaldos.filter((a) => !ids.has(a.obra_id)); }
          res = clon(objetivo);
        }
        guardar(db);
        if (this.uno) {
          if (res.length > 1) return err('JSON object requested, multiple (or no) rows returned', 'PGRST116');
          if (!res.length) return this.uno === 'single' ? err('JSON object requested, multiple (or no) rows returned', 'PGRST116') : { data: null, error: null };
          return { data: res[0], error: null };
        }
        return { data: this.op === 'select' || this.devolver ? res : null, error: null };
      }
    }

    async function rpc(nombre) {
      await espera();
      const db = cargar();
      if (nombre === 'sgo_clave_cambiada') { const p = perfil(db); if (p) p.debe_cambiar_clave = false; guardar(db); return { data: null, error: null }; }
      return err('function not found', '42883');
    }
    const functions = {
      async invoke(nombre, { body }) {
        await espera(40);
        const db = cargar();
        const falla = (msg, status = 400) => ({ data: null, error: { message: 'Edge Function returned a non-2xx status code', context: { status, json: async () => ({ error: msg }) } } });
        if (nombre !== 'sgo-usuarios') return falla('Function not found', 404);
        if (!db.session) return falla('Inicie sesión.', 401);
        if (!esAdmin(db)) return falla('Solo un administrador de SGO puede gestionar usuarios.', 403);
        const b = body || {};
        if (b.accion === 'crear') {
          const email = String(b.email || '').trim().toLowerCase();
          if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return falla('El correo no es válido.');
          if (b.clave && b.clave.length < 8) return falla('La contraseña debe tener al menos 8 caracteres.');
          let u = db.users.find((x) => x.email === email);
          const existia = !!u;
          let temporal = false;
          if (!u) { if (!b.clave) return falla('Indique una contraseña temporal (mínimo 8 caracteres).'); u = { id: 'u-' + db.seq++, email, password: b.clave }; db.users.push(u); temporal = true; }
          else if (b.restablecer) { u.password = b.clave; temporal = true; }
          const previo = db.sgo_usuarios.find((x) => x.id === u.id);
          if (previo) Object.assign(previo, { nombre: b.nombre || '', rol: b.rol === 'admin' ? 'admin' : 'cliente', activo: true, ...(temporal ? { debe_cambiar_clave: true } : {}) });
          else db.sgo_usuarios.push({ id: u.id, email, nombre: b.nombre || '', rol: b.rol === 'admin' ? 'admin' : 'cliente', activo: true, debe_cambiar_clave: temporal, creado: ahora(), creado_por: uid() });
          guardar(db);
          return { data: { ok: true, id: u.id, existia, yaEraUsuario: !!previo, temporal }, error: null };
        }
        if (b.accion === 'clave') {
          if (!b.clave || b.clave.length < 8) return falla('La contraseña debe tener al menos 8 caracteres.');
          const s = db.sgo_usuarios.find((x) => x.id === b.id);
          if (!s) return falla('Ese usuario no pertenece a SGO.', 404);
          db.users.find((x) => x.id === b.id).password = b.clave;
          if (b.id !== uid()) s.debe_cambiar_clave = true;
          guardar(db);
          return { data: { ok: true }, error: null };
        }
        if (b.accion === 'quitar') {
          if (b.id === uid()) return falla('No puede quitarse el acceso a usted mismo.');
          db.sgo_usuarios = db.sgo_usuarios.filter((x) => x.id !== b.id);
          db.sgo_accesos = db.sgo_accesos.filter((x) => x.usuario_id !== b.id);
          guardar(db);
          return { data: { ok: true }, error: null };
        }
        return falla('Acción desconocida.');
      },
    };
    return { auth, from: (t) => new Consulta(t), rpc, functions };
  }
  window.supabase = { createClient };
  window.__falso = { leer: cargar, escribir: guardar, reiniciar: () => localStorage.removeItem(KEY) };
})();
