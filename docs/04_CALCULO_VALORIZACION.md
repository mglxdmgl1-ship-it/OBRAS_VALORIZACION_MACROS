# Cálculo de la valorización mensual

Reglas que usa el libro de valorización (obra por administración directa). La web debe calcular exactamente lo mismo.

## Por partida (cuadro 4.4)

| Concepto | Fórmula |
|---|---|
| Parcial del presupuesto | `metrado_ET × P.U.` |
| Metrado actual (del mes) | Σ metrados diarios del periodo (planillas 4.2 / RESUMEN) |
| Valorizado actual | `metrado_actual × P.U.` (redondeo a 2 decimales) |
| % actual | `valorizado_actual / parcial` |
| Acumulado | `anterior + actual` (metrado y S/) |
| Saldo | `presupuesto − acumulado` (metrado y S/) |
| Títulos (T) | Suma de sus hijas (por `padre`) en cada columna |

Validaciones: el acumulado no debe superar el metrado del expediente (si lo supera → mayor metrado, requiere aprobación); el metrado del día no debería superar el rendimiento del A.C.U. × días; una partida con insumos sin O/S ni ingreso que pesan más del umbral (20 % del costo) **no se valoriza**.

## Resumen (cuadros 4.3 y 4.3.1)

- Subpresupuestos 01…08 = suma de sus partidas.
- **Costo directo (C.D.)** = Σ subpresupuestos.
- Cada costo indirecto se valoriza en proporción al C.D.: `indirecto_valorizado = pct_cd × C.D._valorizado`, con `pct_cd = monto_indirecto / C.D._presupuesto`.

| Concepto (obra 881) | Presupuesto S/ | % del C.D. |
|---|---:|---:|
| Costo directo | 1 063 525.99 | 100.00 % |
| Gastos generales de ejecución | 205 666.44 | 19.34 % |
| Gastos de supervisión | 130 032.94 | 12.23 % |
| Gastos de liquidación | 28 716.67 | 2.70 % |
| Gastos de seguimiento | 13 929.05 | 1.31 % |
| Gastos de monitoreo | 155 892.72 | 14.66 % |
| Elaboración del expediente técnico | 134 208.10 | 12.62 % |
| **Presupuesto total** | **1 731 971.91** | |

- Formato 10 (4.3.1): `Subtotal = C.D. + mayores metrados + partidas nuevas − deductivos`.
- % avance físico = `C.D. acumulado / C.D. presupuesto` (también se informa sobre el presupuesto total).

## Comparación con lo programado (4.5, Curva S, F-12B)

- Programado del mes = cronograma valorizado vigente (4.13) del mes, por partida.
- Diferencia = `ejecutado − programado`; % de cumplimiento = `ejecutado / programado`.
- Curva S: % acumulado programado vs % acumulado físico (C.D.) vs % acumulado financiero (devengado / presupuesto total).
- Valorización N° 01 de la obra 881: programado S/ 60 213.30, ejecutado S/ 54 534.96 → cumplimiento 90.57 %.

## Periodo

- Primer mes: desde la fecha del acta de inicio; último mes: hasta la fecha de término (con ampliaciones).
- Días fuera del plazo no se valorizan. Calendario: lunes a sábado, 8 h/día, con feriados de la hoja BD.
- Al cerrar el mes: actual → anterior; N° de valorización + 1.

## Ejemplo (partida 01.01.01.01.01)

Limpieza, trazo y nivelación: 130.53 m² × S/ 1.97 = S/ 257.14. Metrado del mes 130.53 (días 15, 16 y 21) → valorizado S/ 257.14 (100 %), saldo 0.
