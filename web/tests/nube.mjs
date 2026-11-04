// Prueba en el navegador del inicio de sesión y los roles (Supabase simulado: tests/supabase_falso.js).
// Uso: servir web/public (python -m http.server 8765) y luego
//   node web/tests/nube.mjs http://localhost:8765/ carpeta_salida
import { createRequire } from 'node:module';
import { readFileSync, mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
const { chromium } = createRequire(import.meta.url)('playwright');

const BASE = process.argv[2] || 'http://localhost:8765/';
const OUT = process.argv[3] || './salida_nube';
mkdirSync(OUT, { recursive: true });
const FALSO = readFileSync(fileURLToPath(new URL('./supabase_falso.js', import.meta.url)), 'utf8');
const errores = [];
const pasos = [];
const ok = (cond, msg) => { if (cond) pasos.push('✔ ' + msg); else errores.push('✘ ' + msg); };

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
await ctx.route('**/vendor/supabase.js', (r) => r.fulfill({ body: FALSO, contentType: 'application/javascript' }));
const page = await ctx.newPage();
page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) errores.push('[console] ' + m.text()); });
page.on('pageerror', (e) => errores.push('[pageerror] ' + e.message));
page.on('dialog', (d) => d.accept());
const espera = (ms) => page.waitForTimeout(ms);
const db = () => page.evaluate(() => window.__falso.leer());
async function ingresar(email, clave) {
  await page.waitForSelector('#f-ingreso', { timeout: 15000 });
  await page.fill('#f-ingreso [name=email]', email);
  await page.fill('#f-ingreso [name=clave]', clave);
  await page.click('#ing-btn');
}
async function salir() {
  await page.click('#b-salir');
  await page.waitForSelector('#f-ingreso', { timeout: 15000 });
}
const irA = async (r) => { await page.evaluate((x) => (location.hash = '#/' + x), r); await espera(700); };

// ---------------------------------------------------------------- administrador
await page.goto(BASE);
await page.evaluate(() => window.__falso.reiniciar());
await page.reload();
await page.waitForSelector('#f-ingreso', { timeout: 15000 });
await page.screenshot({ path: `${OUT}/01_ingreso.png` });
ok(await page.isHidden('#app'), 'sin sesión solo se ve la pantalla de ingreso');
await ingresar('admin@prueba.pe', 'clave-mala');
await page.waitForSelector('#ing-error:not([hidden])');
ok((await page.innerText('#ing-error')).includes('Correo o contraseña incorrectos'), 'contraseña incorrecta: mensaje en español');
await ingresar('admin@prueba.pe', 'Admin12345');
await page.waitForSelector('#b-ejemplo', { timeout: 15000 });
ok((await page.innerText('#usuario')).includes('Administrador'), 'administrador ingresa y ve su rol');
ok(await page.isVisible('a[href="#/usuarios"].menu-item'), 'el menú del administrador tiene «Usuarios y accesos»');
await page.click('#b-ejemplo');
await page.waitForSelector('#g-curva', { timeout: 30000 });
await espera(1500);
let d = await db();
ok(d.sgo_obras.length === 1 && d.sgo_obras[0].version === 1, 'la obra de ejemplo se guarda en la nube (versión 1)');
const obraId = d.sgo_obras[0].id;
await page.screenshot({ path: `${OUT}/02_admin_panel.png` });

// edición → subida automática con versión
await irA('presupuesto');
const inp = await page.$('#t-pres tbody tr:not(.tit) input[data-k="metrado"]');
await inp.fill('140');
await inp.press('Tab');
await page.waitForFunction(() => window.__falso.leer().sgo_obras[0]?.version === 2, null, { timeout: 15000 });
d = await db();
ok(d.sgo_obras[0].datos.partidas.find((p) => p.tipo === 'P').metrado === 140, 'el cambio de metrado se sube a la nube (versión 2)');
ok(d.sgo_respaldos.length === 1, 'se guarda el respaldo automático del día');
ok((await page.innerText('#estado-guardado')).includes('En la nube'), 'estado: «En la nube»');

// crear cliente y asignarle la obra
await irA('usuarios');
await page.click('#b-nuevo');
await page.fill('.modal [name=nombre]', 'SUPERVISIÓN DE PRUEBA');
await page.fill('.modal [name=email]', 'cliente@prueba.pe');
await page.fill('.modal [name=clave]', 'Temporal123');
await page.click('.modal-pie .btn-primario');
await page.waitForSelector('#cred', { timeout: 15000 });
ok((await page.inputValue('#cred')).includes('Temporal123'), 'se muestran las credenciales del nuevo cliente');
await page.screenshot({ path: `${OUT}/03_credenciales.png` });
await page.click('.modal-pie button:has-text("Listo")');
await espera(800);
d = await db();
ok(d.sgo_usuarios.some((u) => u.email === 'cliente@prueba.pe' && u.rol === 'cliente' && u.debe_cambiar_clave), 'cliente creado con contraseña temporal');
await page.click('[data-tab="accesos"]');
await espera(700);
await page.check(`#t-acc input[data-obra="${obraId}"]`);
await espera(800);
d = await db();
ok(d.sgo_accesos.length === 1, 'acceso del cliente a la obra registrado');
await page.screenshot({ path: `${OUT}/04_accesos.png` });

// ---------------------------------------------------------------- cliente
await salir();
await ingresar('cliente@prueba.pe', 'Temporal123');
await page.waitForSelector('#f-clave', { timeout: 15000 });
ok(true, 'con contraseña temporal se pide crear una nueva');
await page.fill('#f-clave [name=c1]', 'corta');
await page.fill('#f-clave [name=c2]', 'corta');
await page.click('#cl-btn');
ok(!(await page.isHidden('#cl-error')), 'rechaza contraseñas cortas');
await page.fill('#f-clave [name=c1]', 'MiClaveNueva2026');
await page.fill('#f-clave [name=c2]', 'MiClaveNueva2026');
await page.click('#cl-btn');
await page.waitForSelector('#g-curva', { timeout: 30000 });
await espera(800);
ok((await page.innerText('#usuario')).includes('Cliente'), 'el cliente ingresa a su obra');
ok(!(await page.isVisible('a[href="#/usuarios"].menu-item')) && !(await page.isVisible('a[href="#/calculadora"].menu-item')), 'el cliente no ve Usuarios ni Calculadora');
await page.screenshot({ path: `${OUT}/05_cliente_panel.png` });
await irA('presupuesto');
ok(await page.$eval('#t-pres tbody tr:not(.tit) input[data-k="metrado"]', (e) => e.disabled), 'presupuesto: los campos están bloqueados');
ok(await page.isHidden('#b-imp'), 'presupuesto: no se puede importar');
ok(await page.isVisible('[data-exp="pdf"][data-rep="pres"]'), 'presupuesto: puede exportar a PDF');
const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 60000 }), page.click('[data-exp="xlsx"][data-rep="pres"]')]);
ok(/PRESUPUESTO/.test(dl.suggestedFilename()), 'el cliente descarga el Excel del presupuesto');
await page.screenshot({ path: `${OUT}/06_cliente_presupuesto.png` });
await irA('valorizacion');
ok(await page.isHidden('#b-cerrar'), 'valorización: el cliente no puede cerrar el periodo');
ok(!(await page.$eval('#sel-per', (e) => e.disabled)), 'valorización: puede elegir el periodo');
await irA('acu');
await page.click('#a-lista tr:nth-child(3)');
await espera(500);
ok(await page.$eval('#t-acu tbody input[data-k="cantidad"]', (e) => e.disabled), 'A.C.U.: al cambiar de partida los campos siguen bloqueados');
await irA('calculadora');
ok((await page.innerText('.vista-cab h1')).includes('I.E.I.') || (await page.innerText('.vista-cab h1')).includes('Panel'), 'ruta de calculadora redirige al panel');
const v0 = (await db()).sgo_obras[0].version;
await irA('obra');
ok(await page.isHidden('#b-guardar'), 'datos de la obra: sin botón Guardar');
await espera(3000);
ok((await db()).sgo_obras[0].version === v0, 'el cliente no genera versiones nuevas');
await irA('cuenta');
ok(!(await page.$eval('#f-clave [name=c1]', (e) => e.disabled)), 'Mi cuenta: el cliente puede cambiar su contraseña');
await page.screenshot({ path: `${OUT}/07_cliente_cuenta.png` });

// el administrador guarda una versión nueva → al cliente le llega sola (sondeo cada 60 s; se fuerza aquí)
await page.evaluate(() => { const d = window.__falso.leer(); d.sgo_obras[0].datos.datos.nombre_corto = 'OBRA RENOMBRADA'; d.sgo_obras[0].version++; window.__falso.escribir(d); });
await page.evaluate(async () => { const m = await import('./js/store.js'); await m.revisarVersionRemota(); });
await espera(1200);
ok((await page.innerText('#sel-obra')).includes('OBRA RENOMBRADA') || (await page.innerText('#vista')).includes('OBRA RENOMBRADA'), 'el cliente recibe la versión más reciente');

// ---------------------------------------------------------------- conflicto de versiones (administrador)
await salir();
await ingresar('admin@prueba.pe', 'Admin12345');
await page.waitForSelector('#g-curva', { timeout: 30000 });
await irA('presupuesto');
await page.evaluate(() => { const d = window.__falso.leer(); d.sgo_obras[0].version += 5; window.__falso.escribir(d); });
const inp2 = await page.$('#t-pres tbody tr:not(.tit) input[data-k="metrado"]');
await inp2.fill('150');
await inp2.press('Tab');
await page.waitForSelector('.modal:has-text("Otra persona guardó cambios")', { timeout: 15000 });
ok(true, 'conflicto de versiones detectado y consultado al usuario');
await page.screenshot({ path: `${OUT}/08_conflicto.png` });
await page.click('.modal-pie .btn-primario');
await page.waitForFunction(() => window.__falso.leer().sgo_obras[0].datos.partidas.find((p) => p.tipo === 'P').metrado === 150, null, { timeout: 15000 });
ok(true, '«Conservar mis cambios» sube la versión local');

// desactivar al cliente
await irA('usuarios');
const fila = await page.$('#t-us tbody tr:has-text("cliente@prueba.pe")');
await (await fila.$('[data-k="activo"]')).uncheck();
await espera(900);
ok((await db()).sgo_usuarios.find((u) => u.email === 'cliente@prueba.pe').activo === false, 'el administrador desactiva al cliente');
// no puede quitarse el rol a sí mismo ni dejar el sistema sin administrador
ok(await page.$eval('#t-us tbody tr:has-text("admin@prueba.pe") [data-k="rol"]', (e) => e.disabled), 'el administrador no puede cambiar su propio rol');
await salir();
await ingresar('cliente@prueba.pe', 'MiClaveNueva2026');
await page.waitForSelector('#sa-salir', { timeout: 15000 });
ok((await page.innerText('#acceso')).includes('desactivada'), 'cliente desactivado: sin acceso');
await page.screenshot({ path: `${OUT}/09_sin_acceso.png` });

// móvil: pantalla de ingreso
await page.click('#sa-salir');
await page.setViewportSize({ width: 390, height: 844 });
await page.waitForSelector('#f-ingreso');
await page.screenshot({ path: `${OUT}/10_ingreso_movil.png` });

console.log(pasos.join('\n'));
console.log(JSON.stringify({ pasos: pasos.length, errores }, null, 1));
await browser.close();
process.exit(errores.length ? 1 : 0);
