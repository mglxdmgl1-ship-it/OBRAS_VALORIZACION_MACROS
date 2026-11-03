# Expediente técnico revisado (no incluido en el repositorio)

Carpeta: `05_OBRAS\01_ET_INSTITUCION_EDUCATIVA` (~3.8 GB, 737 archivos: PDF, planos, Excel). No se sube a GitHub por tamaño; aquí queda el inventario de lo que alimenta a los programas.

## Estructura del expediente

| Tomo | Contenido | Uso en los programas |
|---|---|---|
| I. Resumen ejecutivo | 1.1 | Datos generales (BD). |
| II. Memoria descriptiva | 2.1 general … 2.11 costos y presupuestos | Coordenadas, altitud, área (BD). |
| III. Especificaciones técnicas | 3.1 obras provisionales … 3.8 mobiliario | Método de medición y forma de pago de cada partida (planillas 4.2). |
| IV. Estudios básicos | Mecánica de suelos, topografía, riesgos, SSOMA, demanda, demolición, impacto ambiental, EVAR, CIRA, georreferenciación, plan de contingencia, obras provisionales | Referencia. |
| V. Ingeniería del proyecto | Memorias de cálculo (estructuras, sanitarias, eléctricas, mecánicas, comunicaciones, flete, señalética) | Referencia. |
| VI. Metrados | 6.1 resumen y 6.2 planillas por especialidad (Excel y PDF) | Metrado del E.T. por partida; detalle por elemento/eje para las planillas diarias. |
| **VII. Presupuesto de obra** | 7.2 resumen, 7.3 presupuesto consolidado y desagregado, **7.4 A.C.U.**, 7.5 relación de insumos, **7.6 presupuesto analítico**, 7.7 costo indirecto, 7.8 base de datos (S10), 7.9 cotizaciones | **Se importa** en el control de metrados (PRESUPUESTO, ACU); P.U. e insumos del catálogo; % de costos indirectos. |
| **VIII. Cronogramas** | 8.1 Gantt (MS Project, XML conciliado, insumos CSV), 8.2 adquisiciones, mano de obra y maquinaria, **8.3 cronograma valorizado** | Cronograma valorizado mensual (4.13), programado para 4.5 y Curva S. |
| IX. Fórmula polinómica | 9 | Reajustes (no aplica en adm. directa). |
| X. Planos | 272 archivos | Referencia para metrados. |
| XI. Otros / XII. Anexos | Varios | Referencia. |
| Expediente escaneado | 13 tomos PDF | Copia firmada. |

## Otras carpetas de `05_OBRAS` relacionadas

| Carpeta | Contenido |
|---|---|
| `04_INFORMES\07_INFORMES_MENSUALES` | Ejemplos de informes y el archivo digital del informe mensual N° 01 (octubre 2026). |
| `04_INFORMES\03_REQUERIMIENTOS` | Requerimientos (O/C, O/S) que alimentan los ingresos de almacén. |
| `04_INFORMES\04_CRONOGRAMA_SEGUIMIENTO`, `08_INFORMES_DE_CORTE`, `09_INFORMES_PARA_CONSEJERO` | Seguimiento y reportes. |
| `05_PIM`, `06_CMN`, `07_DATOS` | Modificación de PIM, cuadro multianual de necesidades, resolución de aprobación, disponibilidad presupuestal, presupuesto analítico escaneado. |
| `03_NORMATIVAS\webapp` | Aplicación web existente **SGO-IEI881** (Node.js + Express + ExcelJS + Multer): procedimientos de la Ley N° 32069, cronograma, calendario y expediente técnico con importación desde Excel. Base posible para la página de valorizaciones. |

## Datos clave de la obra 881

- Proyecto: «Mejoramiento del servicio de educación inicial en I.E. 881, distrito de Cabanilla, provincia de Lampa, Puno». C.P. Miraflores, 3 882 m s.n.m., terreno 1 970 m².
- CUI 2649805 · Meta 0180 · Rubro 15 FONCOR · Función 22 Educación / 047 / 0103.
- E.T. aprobado con R.G.R. N° 177-2026-GRI-GR PUNO (07/04/2026): S/ 1 731 971.91 (C.D. S/ 1 063 525.99). Modalidad adm. directa: R.G.R. N° 393-2026-GRI-GR PUNO.
- Inicio 15/10/2026 · 120 d.c. · término 11/02/2027 · 5 meses de ejecución (oct-2026 a feb-2027).
- 8 subpresupuestos: 01 obras provisionales y SST, 02 estructuras, 03 arquitectura, 04 sanitarias, 05 eléctricas, 06 telecomunicaciones, 07 mecánicas, 08 mobiliario y equipamiento.
