# 2. Almacén y equipo mecánico

Lleva el **almacén de obra** y el **control de la maquinaria** en un solo libro y arma los formatos del informe mensual: kardex, bincard, partes diarios y los cuadros **4.10** (movimiento diario de almacén), **4.11** (ingresos – egresos), **4.12** (saldo valorizado) y **4.15** (control de equipo mecánico). Puede importar del *Control de metrados diarios* los ingresos, las salidas calculadas (A.C.U. × metrado del día) y los partes de los equipos.

| Archivo | Uso |
|---|---|
| `PROGRAMA_ALMACEN_EQUIPO_v008.xlsm` | Versión vigente. |
| `GUIA_USO_PROGRAMA_ALMACEN.xlsx` | Dónde se llena cada cosa y orden de uso mes a mes (con casilla HECHO). |
| `src/*.bas` | Código VBA v008. El historial de git tiene v002 → v008 (`git log -- programas/02_almacen_equipo_mecanico/src`). |

## Hojas

| Hoja | Se digita / se calcula | Contenido |
|---|---|---|
| INICIO, AYUDA | — | Menú (pasos 1 a 4) y ayuda. |
| DATOS | Se digita | Entidad, gerencia, proyecto, CUI, meta, ubicación, modalidad, mes que se controla, fecha de corte, N° de informe, 4 cargos que firman, carpetas y rutas (control de metrados, Excel de informes 4.10 y 4.15), opciones de importación. |
| INV. 2026 | Se digita | Una fila por ingreso: ítem, descripción, unidad, tipo de documento, P.U., cantidad, total, fecha, N° req., proveedor. Las filas sin unidad ni cantidad son los RUBROS. |
| MOV. ALMACEN | Se digita | Saldo anterior y cantidades de ENTRADA/SALIDA de cada día (cuadro 4.10). |
| DETALLE SALIDAS | Se completa | Por salida: N° de vale/PECOSA, pedido por, destino/partida. |
| INGRESOS - EGRESOS / SALDO VALORIZADO | Se calcula | Cuadros 4.11 y 4.12 con fórmulas. |
| EQUIPOS | Se digita | Por máquina: código (E-01…), descripción, marca, modelo, serie/placa, proveedor, O/S, unidad (H.M./H.E./día), cantidad contratada, costo unitario, condición (seca/con combustible), operador; calcula acumulado, saldo de la O/S y valorizado del mes. |
| MOV. EQUIPOS | Se digita | Por equipo y día: operador, horario mañana/tarde, horómetro inicial/final, horas, combustible (gal), partida, trabajo. |
| CONTROL EQUIPO MECANICO | Se calcula | Cuadro 4.15: horas y galones por día. |
| AUTOLLENADO | Opcional | Jornada por defecto, trabajo por defecto, horómetro inicial, pedido por (maestro de obra o por rubro). |
| REVISION | Se calcula | Lista de lo que falta o no cuadra (hoja, fila, dato, observación). |
| CATALOGO / MOVIMIENTOS / PARTES | Base interna | Catálogo de insumos (código M-01…, P.U., saldos), movimientos (ingreso/salida con documento, O/C-O/S, partida, saldo después), partes diarios (horas reales, según A.C.U., a valorizar, combustible). |
| KARDEX / BINCARD / PARTE_DIARIO | Se arma | Formatos listos para imprimir (desde FMT_KARDEX, FMT_BINCARD, FMT_PARTE). |
| FMT_* , LISTAS | Plantillas | Formatos base del informe y listas desplegables. |

## Botones (macros)

| Botón | Macro | Módulo |
|---|---|---|
| ④ ACTUALIZAR CUADROS | `ActualizarCuadrosBoton` | ModReporte |
| ACTUALIZAR Y REVISAR | `ActualizarYRevisar` | ModEquipos |
| GENERAR KARDEX / BINCARD / PARTES DIARIOS | `GenerarKardex` / `GenerarBincard` / `GenerarPartes` | ModKardex / ModPartes |
| GENERAR LOS TRES | `GenerarTodo` | ModMenu |
| PDF KARDEX (A4 horizontal) / PDF BINCARD (A5 o A4 doble) / PDF PARTES DIARIOS / PDF DE TODO | `PdfKardex` / `PdfBincard` / `PdfPartes` / `PdfTodo` | ModPDF |
| INFORME DE ALMACÉN / INFORME DE EQUIPO MECÁNICO | `ExportarInformeAlmacenBoton` / `ExportarInformeEquiposBoton` | ModReporte |
| AUTOLLENADO / DETALLE DE SALIDAS | `IrAutollenado` / `IrDetalle` | ModAuto |
| TRAER DATOS ANTERIORES | `TraerDatosAnteriores` | ModReporte |
| PROCESAR Y VALIDAR | `ProcesarTodo` | ModProcesar |
| CARPETA DE LOS PDF · CERRAR MES · NUEVA OBRA | `AbrirCarpeta` · `CerrarMes` · `NuevaObra` | ModMenu |
| LOGO DE LOS FORMATOS | `CambiarLogo` | ModBase |

## Módulos

| Módulo | Líneas | Qué hace |
|---|---|---|
| `ModBase.bas` | 816 | Constantes (celdas de DATOS, filas y columnas), utilidades de texto/número/fecha, carpetas y PDF, logo, modo prueba. |
| `ModImportar.bas` | 791 | Importa del control de metrados: ingresos, salidas por A.C.U. × metrado, partes de equipos; agrega al catálogo materiales y equipos nuevos; excluye insumos por palabra. |
| `ModReporte.bas` | 1274 | Lee INV. y EQUIPOS, arma los cuadros 4.10, 4.11, 4.12 y 4.15 sin perder lo digitado, lista salidas, exporta los informes de almacén y equipo. |
| `ModKardex.bas` | 428 | Lee almacén y catálogo, arma kardex y bincard (área B:P, logo, recuadros). |
| `ModPartes.bas` | 466 | Lee equipos y partes, calcula horas por tramo, arma los partes diarios. |
| `ModEquipos.bas` | 464 | Movimiento diario de equipos → control de equipo y partes; hoja REVISIÓN. |
| `ModAuto.bas` | 448 | Detalle de salidas (vale, pedido por, destino) y autollenado de partes. |
| `ModProcesar.bas` | 167 | Saldo después de cada movimiento, saldos negativos, control de horómetro/horario y saldo de O/S. |
| `ModPDF.bas` | 514 | Escala cada página al papel; A5 o A4 con dos por cara (original y copia). |
| `ModInforme.bas` | 258 | Envía los datos a las hojas `BD_` de un Excel de informe externo (4.10 y 4.15). |
| `ModMenu.bas` | 151 | Generar todo, abrir carpeta, respaldo, cerrar mes, nueva obra. |

## Orden de uso (mes a mes)

1. DATOS (una vez por obra) → 2. INV. y EQUIPOS → 3. ACTUALIZAR CUADROS → 4. Registrar cada día (MOV. ALMACÉN y MOV. EQUIPOS) → 5. DETALLE SALIDAS → 6. ACTUALIZAR Y REVISAR → 7. Generar formatos → 8. PDF → 9. CERRAR MES.

Los PDF se guardan en `SALIDAS_AAAA_MM/` (KARDEX, BINCARD, INFORME) junto al programa (carpeta ignorada en git).

## Historial de versiones

| Versión | Cambio principal |
|---|---|
| v002 | Base: catálogo, movimientos, kardex/bincard, partes, importación, envío al informe. |
| v003 | `ModReporte`: cuadros del informe dentro del programa. |
| v004 | Ajustes de cuadros y menú. |
| v005 | `ModPDF`: impresión A5 / A4 doble. |
| v006 | `ModAuto`: detalle de salidas y autollenado. |
| v007 | `ModEquipos`: movimiento diario de equipos y REVISIÓN. |
| v008 | Vigente (ajustes finales; igual a `VBA_FUENTES`). |
