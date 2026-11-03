# SGO – Costos, presupuestos y valorizaciones (página web)

Aplicación web que reúne en un solo programa la lógica de los tres libros de Excel/VBA del repositorio
(control de metrados e insumos, almacén y equipo mecánico, valorización mensual). Sirve para **cualquier obra**:
se puede empezar de cero llenando todo a mano o importar desde Excel (S10, Delphin, Kennedy, los propios libros de macros) o JSON.

Es una página estática (HTML + JavaScript, sin servidor): se publica en Netlify desde `web/public` (ver `netlify.toml` en la raíz)
y también funciona abriéndola con cualquier servidor local. Los datos se guardan en el navegador (IndexedDB); use **Respaldo**
para descargar la obra completa en un `.json` y llevarla a otra computadora.

## Módulos

| Pantalla | Qué hace | Equivale a |
|---|---|---|
| Panel | Presupuesto, valorizado, % físico, % programado, cumplimiento, plazo, curva S y alertas | — |
| Datos de la obra | Identificación, plazo y ampliaciones, responsables y firmas, costos indirectos, documentos, feriados, umbral «no valorizar», redondeo | Hojas DATOS / BD |
| Presupuesto | Árbol de títulos y partidas, edición, importación de Excel (con subpresupuestos), P.U. del A.C.U. | PRESUPUESTO |
| A.C.U. | Insumos por partida (MO, MAT, EQ, SC, %MO, SP), cuadrilla y rendimiento, copiar A.C.U., importar formato S10 por bloques o tabla plana | ACU, `ModImportar` |
| Insumos y precios | Catálogo único de precios (cambia todos los A.C.U.), relación total de insumos del presupuesto | INSUMOS |
| Calculadora de insumos | Metrado libre × A.C.U. → cantidades; redondeo a enteros (o múltiplos) y **cantidad final editable**; registrar salida de almacén o metrado del día | — |
| Metrados diarios | Planilla del mes (RESUMEN), planilla del día con detalle (N° veces × largo × ancho × alto), estados VÁLIDO / REVISAR / SUPERA RENDIMIENTO / SUPERA METRADO / NO VALORIZAR / FUERA DEL PLAZO, insumos del mes vs ingresos | RESUMEN, 01…31, INSUMOS, `ModActualizar.Calcular` |
| Almacén | Catálogo enlazado al A.C.U. (con factor, p. ej. varilla → kg), ingresos y salidas, saldo después de cada movimiento, **salidas desde A.C.U. × metrado (editables)**, kardex, bincard, vales, cuadros 4.10, 4.11 y 4.12 | Programa de almacén v008 |
| Equipo mecánico | Equipos y O/S, partes diarios (horario, horómetro, combustible), partes desde A.C.U. × metrado, control 4.15 | EQUIPOS, PARTES, `ModPartes` |
| Valorización mensual | Periodos, 4.4 por partida (metrado del mes = Σ diarios o digitado), 4.3 resumen con indirectos, Formato 10, 4.5 comparado, curva S con devengado, cierre del periodo (actual → anterior) | Libro de valorización |
| Cronograma valorizado | 4.13 por partida y mes (S/ o %), generación automática, importación | 4.13 |
| Reportes | Todos los cuadros en **PDF** (A4, encabezado de la obra y firmas) y **Excel con fórmulas** | — |
| Obras, importar y respaldo | Varias obras, obra de ejemplo I.E.I. 881, respaldo/restauración JSON, JSON de la valorización | — |

## Pruebas

```bash
node --test web/tests/*.test.mjs
```

Comprueban los criterios de `prompts/PROMPT_PAGINA_WEB.md` con los datos reales: C.D. S/ 1 063 525.99, total S/ 1 731 971.91,
valorización N° 01 S/ 54 534.96 (5.13 %), indirectos S/ 34 276.24, total del mes S/ 88 811.20, subpresupuestos 01 y 02,
programado S/ 60 213.30 → 90.57 %, curva S 5.66 / 28.97 / 62.79 / 95.73 / 100 %, las 900 filas del cuadro 4.4 contra el Excel,
estados del control diario, cierre de periodo, almacén y los importadores (S10 con subpresupuestos y subpartidas).

Prueba en navegador (recorre todas las pantallas, exporta los 40 PDF/Excel, crea una obra desde cero importando Excel y cierra una valorización):

```bash
cd web/public && python3 -m http.server 8765
```

```bash
node web/tests/navegador.mjs http://localhost:8765/ salida_prueba
```

Las fórmulas de los Excel exportados se verificaron recalculándolos en LibreOffice: dan los mismos valores que la página.

## Datos de ejemplo

`public/datos/iei881.json` se genera con:

```bash
python herramientas/exportar_obra_web.py datos/iei881/valorizacion_01_2026-10.json programas/01_control_metrados_insumos/obra_iei881/CONTROL_METRADOS_DIARIOS_INSUMOS_IEI881_v003.xlsm programas/02_almacen_equipo_mecanico/PROGRAMA_ALMACEN_EQUIPO_v008.xlsm web/public/datos/iei881.json
```

No incluye DNI, celulares ni N° de CIP.

## Estructura

```
web/public/index.html        página
web/public/css/app.css       estilos
web/public/js/util.js        redondeo como Excel, fechas, formatos
web/public/js/calc.js        motor de cálculo (presupuesto, A.C.U., estados, valorización, almacén, equipos)
web/public/js/importar.js    importadores Excel/CSV/JSON
web/public/js/reportes.js    definición de cada cuadro
web/public/js/exportar.js    PDF (jsPDF) y Excel con fórmulas (ExcelJS)
web/public/js/store.js       guardado en IndexedDB
web/public/js/vistas/*.js    pantallas
web/public/vendor/           ExcelJS 4.4.0, jsPDF 2.5.2, AutoTable 3.8.4, Chart.js 4.4.6
web/tests/                   pruebas
```
