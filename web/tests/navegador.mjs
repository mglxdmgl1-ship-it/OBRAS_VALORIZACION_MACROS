// Prueba de humo en el navegador (Playwright): carga la obra de ejemplo, recorre todas las pantallas,
// exporta PDF y Excel y verifica que no haya errores en la consola.
// Uso: servir web/public (p. ej. python -m http.server 8765) y luego
//   node web/tests/navegador.mjs http://localhost:8765/ carpeta_salida
import { createRequire } from 'node:module';
const { chromium } = createRequire(import.meta.url)('playwright');
import { mkdirSync } from 'node:fs';

const URL = process.argv[2] || 'http://localhost:8765/';
const OUT = process.argv[3] || './salida_prueba';
mkdirSync(OUT, { recursive: true });
const errores = [];
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
const page = await ctx.newPage();
page.on('console', (m) => { if (m.type() === 'error') errores.push('[console] ' + m.text()); });
page.on('pageerror', (e) => errores.push('[pageerror] ' + e.message));
page.on('dialog', (d) => d.accept());

const espera = (ms) => page.waitForTimeout(ms);
async function ir(ruta) {
  await page.evaluate((r) => (location.hash = '#/' + r), ruta);
  await espera(700);
  const err = await page.$('#vista > .nota-error');
  if (err) errores.push(`[${ruta}] ` + (await err.innerText()));
}
async function descargar(sel, nombre) {
  const [d] = await Promise.all([page.waitForEvent('download', { timeout: 60000 }), page.click(sel)]);
  const p = `${OUT}/${nombre}_${d.suggestedFilename()}`;
  await d.saveAs(p);
  return p;
}

await page.goto(URL);
await espera(800);
await page.click('#b-ejemplo');
await page.waitForSelector('#g-curva', { timeout: 20000 });
await espera(800);
await page.screenshot({ path: `${OUT}/01_panel.png`, fullPage: false });

const rutas = ['obra', 'presupuesto', 'acu', 'insumos', 'calculadora', 'metrados', 'almacen', 'equipos', 'valorizacion', 'cronograma', 'reportes', 'obras'];
for (const [i, r] of rutas.entries()) {
  await ir(r);
  await page.screenshot({ path: `${OUT}/${String(i + 2).padStart(2, '0')}_${r}.png` });
}

// pestañas
for (const t of ['dia', 'insumos', 'mes']) { await ir('metrados'); await page.click(`[data-tab="${t}"]`); await espera(600); await page.screenshot({ path: `${OUT}/metrados_${t}.png` }); }
for (const t of ['movimientos', 'generar', 'kardex', 'cuadros', 'stock']) { await ir('almacen'); await page.click(`[data-tab="${t}"]`); await espera(600); await page.screenshot({ path: `${OUT}/almacen_${t}.png` }); }
for (const t of ['partes', 'generar', 'control', 'equipos']) { await ir('equipos'); await page.click(`[data-tab="${t}"]`); await espera(500); await page.screenshot({ path: `${OUT}/equipos_${t}.png` }); }
for (const t of ['v43', 'v45', 'curva', 'ant', 'v44']) { await ir('valorizacion'); await page.click(`[data-tab="${t}"]`); await espera(600); await page.screenshot({ path: `${OUT}/valorizacion_${t}.png` }); }

// calculadora: agregar partida con metrado y editar cantidad
await page.evaluate(() => (location.hash = '#/calculadora?item=02.01.03.01'));
await espera(800);
const inp = await page.$('#t-lin input[data-k=metrado]');
await inp.fill('37.5');
await inp.press('Tab');
await espera(600);
const fin = await page.$('#t-ins input[data-final]');
if (fin) { await fin.fill('99'); await fin.press('Tab'); await espera(500); }
await page.screenshot({ path: `${OUT}/calculadora_resultado.png` });
const archivos = [];
archivos.push(await descargar('[data-rep="calc"][data-exp="xlsx"]', 'calc'));
archivos.push(await descargar('[data-rep="calc"][data-exp="pdf"]', 'calc'));

// metrado diario: escribir en la grilla
await ir('metrados');
await page.click('[data-tab="mes"]');
await espera(500);
const celda = await page.$('#t-mes tbody tr input[data-f="2026-10-31"]');
if (celda) { await celda.fill('3.25'); await celda.press('Enter'); await espera(600); }

// centro de reportes: todos los formatos
await ir('reportes');
const ids = await page.$$eval('[data-exp]', (bs) => bs.map((b) => b.dataset.rep + '|' + b.dataset.exp));
for (const id of ids) {
  const [rep, exp] = id.split('|');
  try { archivos.push(await descargar(`[data-rep="${rep}"][data-exp="${exp}"]`, 'rep')); }
  catch (e) { errores.push(`[export ${id}] ${e.message}`); }
  await espera(150);
}
const avisosError = await page.$$eval('.aviso-error', (a) => a.map((x) => x.innerText));
errores.push(...avisosError.map((t) => '[aviso] ' + t));

// nueva obra en blanco
await ir('obras');
await page.click('#b-nueva');
await page.fill('[name=nombre]', 'OBRA DE PRUEBA');
await page.fill('[name=nombre_corto]', 'PRUEBA');
await page.click('.modal-pie .btn-primario');
await espera(800);
await page.screenshot({ path: `${OUT}/obra_nueva_presupuesto.png` });

// obra desde cero: importar presupuesto y A.C.U. exportados antes (como si vinieran de S10)
const presX = archivos.find((f) => /PRESUPUESTO_.*\.xlsx$/.test(f));
const acuX = archivos.find((f) => /ANALISIS_COSTOS_UNITARIOS_.*\.xlsx$/.test(f));
await ir('presupuesto');
let [fc] = await Promise.all([page.waitForEvent('filechooser'), page.click('#b-imp')]);
await fc.setFiles(presX);
await page.waitForSelector('.modal-pie .btn-primario', { timeout: 30000 });
await page.screenshot({ path: `${OUT}/importar_presupuesto_previa.png` });
await page.click('.modal-pie .btn-primario');
await espera(1500);
const cdTxt = await page.$eval('.vista-cab p', (e) => e.innerText);
if (!cdTxt.includes('1,063,525.99')) errores.push('[importar presupuesto] C.D. inesperado: ' + cdTxt);
await ir('acu');
[fc] = await Promise.all([page.waitForEvent('filechooser'), page.click('#b-imp')]);
await fc.setFiles(acuX);
await page.waitForSelector('.modal-pie .btn-primario', { timeout: 60000 });
await page.click('.modal-pie .btn-primario');
await espera(2500);
const acuTxt = await page.$eval('.vista-cab p', (e) => e.innerText);
if (!acuTxt.startsWith('653 de 653')) errores.push('[importar A.C.U.] ' + acuTxt);
// valorización de la obra nueva: metrado manual en 4.4 y cierre
await ir('valorizacion');
await page.click('#f-av');
await espera(800);
const act = await page.$('#t-val input[data-act]');
await act.fill('1');
await act.press('Tab');
await espera(800);
await page.click('#b-cerrar');
await page.click('.modal-pie .btn-oro');
await espera(1200);
const titulo = await page.$eval('.vista-cab p', (e) => e.innerText);
if (!titulo.includes('N° 02')) errores.push('[cierre] no se creó la valorización N° 02: ' + titulo);
await page.screenshot({ path: `${OUT}/valorizacion_02.png` });

// vista móvil
await page.setViewportSize({ width: 390, height: 844 });
await ir('panel');
await page.screenshot({ path: `${OUT}/movil_panel.png` });
await ir('calculadora');
await page.screenshot({ path: `${OUT}/movil_calculadora.png` });

console.log(JSON.stringify({ archivos: archivos.length, errores }, null, 1));
await browser.close();
process.exit(errores.length ? 1 : 0);
