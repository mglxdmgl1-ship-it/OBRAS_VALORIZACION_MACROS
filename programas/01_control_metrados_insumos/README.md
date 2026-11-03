# 1. Control de metrados diarios e insumos

Controla **cada día lo que se puede valorizar**: el metrado ejecutado por partida, los insumos que ese metrado consume según el A.C.U. frente a lo que realmente ingresó al almacén (O/C, O/S, actas de préstamo), las horas de maquinaria y combustible, y las HH de mano de obra frente al tareo. Sirve para cualquier obra: se carga el presupuesto y los A.C.U. exportados de S10, Delphin, Kennedy u otro programa.

| Archivo | Uso |
|---|---|
| `CONTROL_METRADOS_DIARIOS_PLANTILLA_GENERAL_v001.xlsm` | Plantilla vacía para una obra nueva. |
| `obra_iei881/CONTROL_METRADOS_DIARIOS_INSUMOS_IEI881_v003.xlsm` | La misma plantilla con la obra I.E.I. 881 cargada (presupuesto 907 filas, A.C.U. ~3 600 filas, ingresos de octubre 2026). |
| `src/*.bas` | Código VBA (idéntico en ambos libros). |

## Hojas

| Hoja | Se digita / se calcula | Contenido |
|---|---|---|
| INICIO, AYUDA | — | Menú con botones y guía de uso. |
| DATOS | Se digita | Entidad, proyecto, CUI, meta, ubicación, fecha de inicio y plazo, mes que se controla, N° de valorización, parámetros (umbral «no valorizar» 20 %, jornada 8 h, uso máximo de HH del tareo 85 %, insumos que no pasan por almacén, palabras de combustibles), cargos que firman. |
| PRESUPUESTO | Importado | ITEM, descripción, und, metrado, P.U., parcial, rendimiento (und/día), tipo (T título / P partida). |
| ACU | Importado | Por partida: insumo, und, tipo (MO/MAT/EQ/SC), cantidad por unidad, precio, parcial, % del P.U. controlado, clase, disponible/confirmado/vigente desde, consumo meses anteriores / del mes / acumulado. |
| RESUMEN | **Se digita el metrado de cada día** (columnas 1…31) | Por partida: metrado E.T., P.U., presupuesto, acumulado anterior, días del mes, metrado del mes, acumulado, % avance; filas de cabecera con valorizado del día, HH A.C.U./HH tareo y alertas. |
| INGRESOS | Se digita | Una fila por ítem de cada documento: fecha, tipo y N° de documento, proveedor, descripción, und y cantidad del documento, insumo del expediente, factor de conversión, cantidad en und. del expediente, partida que cubre, fin de la O/S, estado. |
| INSUMOS | Se calcula | Estado de cada insumo con consumo: precio, controlar, primer ingreso, ingresado, consumo por día y saldo. Rojo = se consume sin ingreso. |
| MAQ_COMB | Se calcula + horas del parte | Equipos (H-M / H-E) y combustible: según A.C.U. vs contratado/ingresado, por día, total del mes, acumulado, saldo O/S, valorizable. |
| MANO_OBRA | Se calcula + tareo | HH según A.C.U. por categoría frente al N° de trabajadores del tareo × jornada. |
| 01 … 31 | Se arma sola | Planilla de metrados de cada día (fecha, día de obra, partidas con metrado, acumulado, P.U., parcial, estado/alerta) y firmas. |
| AUX | Interna | Datos auxiliares. |

## Botones (macros)

| Botón | Macro | Módulo |
|---|---|---|
| ② IMPORTAR PRESUPUESTO | `ImportarPresupuesto` | ModImportar |
| ③ IMPORTAR A.C.U. | `ImportarACU` | ModImportar |
| ④ PREPARAR OBRA | `PrepararObra` | ModPreparar |
| ACTUALIZAR TODO | `ActualizarTodo` | ModActualizar |
| VER PLANILLA DEL DÍA | `PlanillaDelDia` | ModActualizar |
| EXPORTAR PLANILLAS A PDF | `ExportarPlanillasPDF` | ModMenu |
| CERRAR MES | `CerrarMes` | ModMenu |
| NUEVA OBRA | `NuevaObra` | ModMenu |

Existen versiones `...Silencioso` de importar/preparar/actualizar/cerrar para llamarlas desde otro programa sin mensajes.

## Módulos

| Módulo | Líneas | Qué hace |
|---|---|---|
| `ModUtil.bas` | 423 | Constantes de hojas y columnas, lectura segura de celdas (`Num`, `Txt`, `ParseNum`), claves de partida, niveles y colores, utilidades. |
| `ModImportar.bas` | 1201 | Lee Excel de S10/Delphin/Kennedy: detecta títulos de columnas y encabezados de partida («Partida 01.01.01» o «01.01.01 DESCRIPCIÓN»), rendimiento, secciones (MO, materiales, equipos, subcontratos, subpartidas) e insumos; reemplaza subpartidas por sus insumos. |
| `ModPreparar.bas` | 917 | Arma RESUMEN, INSUMOS, MAQ_COMB y MANO_OBRA desde PRESUPUESTO y ACU conservando lo ya digitado (busca por ítem, insumo o categoría). |
| `ModActualizar.bas` | 907 | Calcula consumos = metrado del día × cantidad A.C.U., compara con ingresos, horas de maquinaria y combustible, HH; escribe estados y arma las planillas diarias. |
| `ModMenu.bas` | 273 | Navegación, PDF de planillas, cierre de mes (copia de respaldo, pasa el mes al acumulado anterior, aumenta el N° de valorización), nueva obra. |

## Estados de una partida en el día

`VÁLIDO` · `REVISAR / POR CONFIRMAR` (falta un insumo menor o está por confirmar) · `SUPERA RENDIMIENTO` (metrado del día > rendimiento del A.C.U.) · `SUPERA METRADO DEL E.T.` (acumulado > metrado del expediente) · `PENDIENTE: PRESIONE ACTUALIZAR` · `NO VALORIZAR` (los insumos sin O/S ni ingreso pesan más que el umbral, 20 % del costo por defecto).

## Orden de uso

1. DATOS → 2. Importar presupuesto → 3. Importar A.C.U. → 4. Preparar obra.
2. Cada día: registrar INGRESOS, digitar metrados en RESUMEN, ACTUALIZAR TODO, revisar la planilla del día.
3. Fin de mes: EXPORTAR PLANILLAS A PDF → CERRAR MES.

> No ordene (A→Z) las filas del RESUMEN ni del ACU: el programa ubica los datos por posición y clave.
