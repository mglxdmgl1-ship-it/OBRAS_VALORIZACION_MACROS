# 3. Valorización mensual e informe

Libro sin macros (fórmulas) con la **valorización mensual** y los cuadros del informe mensual de obra por administración directa. Es el modelo que la página web debe reproducir y generalizar.

| Archivo | Uso |
|---|---|
| `plantillas/VALORIZACION_MENSUAL_IEI881_2026-10_v3.xlsx` | Valorización N° 01 (15–31/10/2026) de la I.E.I. 881. Original: `01. VALORIZACIÓN MENSUAL I.E.I. N° 881_OCTUBRE 2026 v3.xlsx`. |
| `referencias/REPORTE_FINANCIERO_EJEMPLO_ROSASPATA_2026-09.xlsx` | Ejemplo de otra obra (setiembre 2026) para los cuadros financieros 4.8 a 4.13 y el resumen de ejecución por adicionales. Original: `FINACIERO AL MES DE setiembre 2026- ROSASPATA.xlsx`. |

## Hojas de la valorización

| Hoja | Cuadro del informe | Contenido |
|---|---|---|
| BD | — | Base de datos: rótulo, meta, CUI, unidad ejecutora, clasificador funcional, ubicación, resoluciones (E.T., modalidad, analítico, viabilidad), plazo y ampliaciones, responsables, N° de informe y periodo, presupuesto y % de costos indirectos, feriados. **Todo lo demás lee de aquí.** |
| CARÁTULA, INDICE | — | Carátula y contenido del informe (≈80 ítems con SI/NO, observación, responsable). |
| FICHA TECNICA | — | Ficha técnica de control de obra. |
| F-12B | Formato 12-B | Seguimiento: devengado mensual por específica de gasto (PIM, certificado, ene–dic, saldo), ejecución física por factor productivo, cronograma valorizado, avance físico vs financiero. |
| 4.1 Met. P.B | 4.1 | Metrado ejecutado en el mes por partida y por día. |
| 15 … 31 | 4.2 | Planilla de metrados de cada día de obra (elemento, N° veces, dimensiones, parcial). |
| 4.3. ResVal P.B. | 4.3 | Resumen por subpresupuesto (01 a 08) + costos indirectos + % de avance. |
| 4.3.1 ResVal V.M. | 4.3.1 (Formato 10) | Resumen: metrados aprobados, mayores metrados, partidas nuevas, deductivos. |
| 4.4 Val.Pres P.B. | 4.4 | Valorización por partida: presupuesto, anterior, actual, acumulado y saldo (metrado, S/, %), observación. |
| 4.5. COV vs ejec | 4.5 | Proyectado (cronograma) vs ejecutado del mes, diferencia y % de cumplimiento. |
| Curva-S | Histograma | Programado, físico ejecutado y financiero (devengado) por mes, % acumulados. |
| 4.13 Cron.Val. | 4.13 | Cronograma valorizado mensual vigente por partida. |

El cálculo está descrito en [docs/04_CALCULO_VALORIZACION.md](../../docs/04_CALCULO_VALORIZACION.md). Para la web, use el JSON de `datos/iei881/` o regenérelo con `herramientas/exportar_valorizacion_json.py`.
