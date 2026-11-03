# Modelo de datos (para la página web)

Entidades que salen de las hojas de los tres programas. Los nombres de campo están en minúsculas y sin tildes para usarlos directamente en la base de datos o en JSON. Entre corchetes, la hoja de origen.

## Obra y periodo

**obra** [DATOS, BD]
`id, nombre, nombre_corto, cui, meta, entidad, gerencia, sub_gerencia, unidad_ejecutora, pliego_ue_siaf, modalidad (ADMINISTRACION DIRECTA | CONTRATA), fuente_financiamiento, funcion, division_funcional, grupo_funcional, region, provincia, distrito, localidad, coordenadas, altitud, area_terreno, fecha_inicio, plazo_dias, fecha_termino, presupuesto_total`

**documento_obra** [BD «EXPEDIENTES»]
`obra_id, tipo (EXPEDIENTE_TECNICO | MODALIDAD | ANALITICO | ET_MODIFICADO | VIABILIDAD | ACTA_INICIO | AMPLIACION_PLAZO | PARALIZACION), numero/resolucion, fecha, monto, dias`

**responsable** [BD «RESPONSABLES», DATOS firmas]
`obra_id, cargo (RESIDENTE | SUPERVISOR | INSPECTOR | ASISTENTE_TECNICO | ALMACENERO | ADMINISTRADOR | ...), nombre, profesion, cip, designacion, desde, hasta` — datos personales (DNI, celular) solo si se protegen.

**costo_indirecto** [BD, 4.3]
`obra_id, concepto (GG | SUPERVISION | LIQUIDACION | SEGUIMIENTO | MONITOREO | EXPEDIENTE), monto, pct_cd`

**periodo** (= valorización mensual) [DATOS, BD «INFORME»]
`obra_id, numero (01, 02…), anio, mes, del, al, estado (ABIERTO | CERRADO), fecha_cierre`

**feriado / calendario** [BD]
`fecha, descripcion` · jornada: lunes a sábado, 8 h/día.

## Presupuesto (expediente técnico)

**partida** [PRESUPUESTO, 4.4]
`obra_id, item ("02.01.03.01"), padre, nivel, tipo (T título | P partida), descripcion, und, metrado, pu, parcial, rendimiento (und/día), subpresupuesto ("01"…"08"), origen (ET | MAYOR_METRADO | PARTIDA_NUEVA | DEDUCTIVO)`

**acu_insumo** [ACU]
`partida_item, insumo_codigo, descripcion, und, tipo (MO | MAT | EQ | SC | SP), cuadrilla, cantidad (por und), precio, parcial`

**insumo** [INSUMOS, CATALOGO]
`codigo (M-01…), descripcion, und, grupo (MATERIALES | HERRAMIENTAS | COMBUSTIBLE | ...), precio_expediente, controlar (SI/NO), nombre_en_acu, factor (und ACU → und almacén)`

**cronograma_mensual** [4.13]
`partida_item, anio_mes, monto_programado` (+ metrado programado si se tiene).

## Control diario

**metrado_diario** [RESUMEN columnas 1…31, hojas 01…31 / 15…31]
`obra_id, partida_item, fecha, metrado, detalle[] (elemento, n_veces, n_elem, diametro, peso, largo, ancho, alto, parcial), estado (VALIDO | REVISAR | SUPERA_RENDIMIENTO | SUPERA_METRADO_ET | NO_VALORIZAR), observacion`

**tareo** [MANO_OBRA]
`fecha, categoria (OPERARIO | OFICIAL | PEON | ...), n_trabajadores, horas`

## Almacén

**ingreso** [INGRESOS, INV.]
`obra_id, fecha, tipo_documento (O/C | O/S | GUIA | ACTA_PRESTAMO | NEA | ...), numero_documento, proveedor, descripcion_documento, und_doc, cantidad_doc, insumo_codigo, factor, cantidad, pu, partida_que_cubre, fecha_fin_os, estado (CONFIRMADO | POR_CONFIRMAR)`

**movimiento_almacen** [MOVIMIENTOS, MOV. ALMACEN]
`obra_id, fecha, tipo (INGRESO | SALIDA | DEVOLUCION), documento, n_documento, oc_os, procedencia_destino, partida_item, insumo_codigo, cantidad, pedido_por, vale_pecosa, observacion, origen (DIGITADO | A.C.U. | INGRESOS DEL CONTROL), saldo_despues`

Derivados: kardex por insumo (entradas, salidas, saldo, valorizado), bincard, 4.10 movimiento diario, 4.11 ingresos – egresos, 4.12 saldo valorizado.

## Equipo mecánico

**equipo** [EQUIPOS]
`codigo (E-01…), descripcion, marca, modelo, serie_placa, anio_capacidad, proveedor, os_numero, os_fecha, unidad (H.M. | H.E. | DIA), cantidad_contratada, costo_unitario, monto_os, condicion (SECA | CON_COMBUSTIBLE | PROPIO), combustible_codigo, operador`

**parte_diario** [MOV. EQUIPOS, PARTES]
`equipo_codigo, fecha, n_parte, operador, manana_de, manana_a, tarde_de, tarde_a, horometro_inicial, horometro_final, horas_horario, horas_horometro, horas_a_valorizar, combustible_gal, partida_item, trabajo, metrado, rendimiento_acu, horas_acu`

Derivados: control de equipo (4.15), saldo de O/S, valorizado del mes por equipo.

## Valorización (calculada)

**valorizacion_partida** [4.4]
`periodo_id, partida_item, metrado_anterior, monto_anterior, metrado_actual, monto_actual, metrado_acumulado, monto_acumulado, pct_acumulado, metrado_saldo, monto_saldo, observacion`

**valorizacion_resumen** [4.3, 4.3.1, F-12B]
`periodo_id, concepto (subpresupuesto o costo indirecto), presupuesto, anterior, actual, acumulado, saldo, pct`

**financiero** [F-12B, reporte financiero]
`obra_id, anio, especifica_gasto (2.6.2.2.2.3 …), descripcion, pim, certificado, devengado_por_mes[12], acumulado, saldo`

## Formato JSON exportado

`datos/iei881/valorizacion_01_2026-10.json` (generado por `herramientas/exportar_valorizacion_json.py`):

```jsonc
{
  "obra":        { "nombre", "nombre_corto", "meta", "cui", "unidad_ejecutora", "modalidad", "ubicacion": {...}, ... },
  "documentos":  [ { "documento", "resolucion", "fecha", "monto" } ],
  "plazo":       { "inicio", "dias", "termino", "ampliaciones": [], "plazo_total", "termino_reprogramado" },
  "responsables":[ { "cargo", "nombre", "profesion", "cip", "designacion" } ],
  "informe":     { "numero", "mes", "anio", "periodo_del", "periodo_al", "valorizacion_n" },
  "presupuesto": [ { "concepto", "monto", "pct_cd" } ],
  "feriados":    [ { "fecha", "descripcion" } ],
  "valorizacion": {
    "partidas": [ { "tipo": "titulo|partida", "item", "padre", "nivel", "descripcion", "und", "metrado", "pu", "parcial",
                    "anterior": {"metrado","monto","pct"}, "actual": {...}, "acumulado": {...}, "saldo": {...}, "observacion" } ],
    "costos_indirectos": [ { "concepto", "presupuesto", "anterior", "actual", "acumulado", "saldo" } ],
    "resumen_subpresupuestos": [ { "tipo": "subpresupuesto|total", "item", "descripcion", "presupuesto", "actual", "acumulado", "pct_acumulado", "saldo" } ]
  },
  "cronograma_valorizado": { "meses": ["OCTUBRE 2026", ...], "partidas": [ { "item", "parcial", "por_mes": [...] } ] },
  "curva_s": [ { "n", "mes", "programado", "programado_pct_acum", "ejecutado", "ejecutado_pct_acum", "devengado", "devengado_pct_acum" } ]
}
```

Comprobación de la obra 881, valorización N° 01: 653 partidas + 247 títulos; suma de parciales = C.D. S/ 1 063 525.99; valorizado del mes C.D. = S/ 54 534.96 (5.13 %); total con indirectos = S/ 88 811.20.
