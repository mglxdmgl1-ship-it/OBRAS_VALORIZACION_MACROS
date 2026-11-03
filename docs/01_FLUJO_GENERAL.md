# Flujo general de la valorización

## Actores

| Rol | Qué hace en el flujo |
|---|---|
| Residente de obra | Registra metrados diarios, arma la valorización y el informe mensual, firma. |
| Asistente técnico | Planillas de metrados, apoyo en el control diario. |
| Almacenero | Ingresos y salidas del almacén, kardex, bincard, vales/PECOSA. |
| Controlador / operador de equipo | Partes diarios: horario, horómetro, combustible. |
| Administrador / Especialista administrativo | O/C, O/S, planillas de personal, manifiesto de gastos, devengado SIAF. |
| Supervisor (o inspector) de obra | Revisa y da conformidad a la valorización y al informe. |

## Ciclo diario

1. **Ingresos**: llega material o empieza una O/S → se registra en `INGRESOS` (control de metrados) y en `INV.`/`MOV. ALMACÉN` (almacén). Documento: O/C, O/S, guía, acta de préstamo.
2. **Metrado del día**: por cada partida trabajada se digita el metrado en `RESUMEN` (columna del día). La planilla del día (`01`…`31`) se arma sola.
3. **Consumo teórico**: metrado × cantidad unitaria del A.C.U. = consumo de cada insumo, horas de equipo y HH de mano de obra.
4. **Validación**: se compara el consumo con lo ingresado y con el tareo; la partida queda `VÁLIDO`, `REVISAR`, `SUPERA RENDIMIENTO`, `SUPERA METRADO DEL E.T.` o `NO VALORIZAR`.
5. **Salidas y partes**: el almacén registra las salidas (o las importa del control: A.C.U. × metrado) y los partes diarios de cada equipo.

## Ciclo mensual

1. Cierre del periodo (fecha de corte = último día del mes o del plazo).
2. Control de metrados → exporta planillas PDF; los metrados del mes pasan al libro de **valorización** (4.1, 4.2, 4.4).
3. Almacén y equipo → genera kardex, bincard, partes, cuadros 4.10, 4.11, 4.12 y 4.15 (Excel y PDF).
4. Valorización → 4.3 resumen, 4.3.1 Formato 10, 4.4 por partida, 4.5 comparado vs cronograma, Curva S, 4.13 cronograma vigente, F-12B con devengado SIAF.
5. Informe mensual: índice de ~80 ítems (generalidades, sustento documentario, ejecución, valorizaciones 4.1–4.19, informes de especialistas, anexos). Ver [03_CONTENIDO_INFORME_MENSUAL.md](03_CONTENIDO_INFORME_MENSUAL.md).
6. **Cerrar mes** en cada programa: respaldo, el mes pasa a «acumulado anterior», saldos pasan a «saldo anterior», N° de valorización + 1.

## Relación entre archivos

```
DATOS de la obra (igual en los 3 libros: entidad, proyecto, CUI, meta, mes, N° valorización, firmas)

Control de metrados (.xlsm) ──ImportarControl──► Almacén y equipo (.xlsm)
   PRESUPUESTO, ACU, RESUMEN,                      INGRESOS → MOVIMIENTOS (ingreso)
   INGRESOS, MAQ_COMB                              A.C.U. × metrado → MOVIMIENTOS (salida)
                                                   horas A.C.U. → PARTES
        │
        │ metrado del mes por partida
        ▼
Valorización mensual (.xlsx)  ◄── cuadros 4.10 / 4.11 / 4.12 / 4.15 (EnviarInforme / Exportar)
   BD → 4.1, 4.2, 4.3, 4.4, 4.5, Curva S, 4.13, F-12B
        │
        ▼ exportar_valorizacion_json.py
   datos/<obra>/valorizacion_NN_AAAA-MM.json  → página web
```

## Lo que la página web debe resolver

Hoy los tres libros repiten los datos de la obra y se conectan copiando rutas en `DATOS`. La web debe tener **una sola base de datos por obra** (presupuesto, A.C.U., metrados diarios, ingresos, salidas, equipos, valorizaciones) y generar desde ahí todos los cuadros. Ver [02_MODELO_DE_DATOS.md](02_MODELO_DE_DATOS.md) y [../prompts/PROMPT_PAGINA_WEB.md](../prompts/PROMPT_PAGINA_WEB.md).
