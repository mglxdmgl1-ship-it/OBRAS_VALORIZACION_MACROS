# Prompt: página web de valorización de obras por administración directa

> Copie todo lo que está debajo de la línea en Claude (o Claude Code) con este repositorio abierto. Está escrito para que el resultado sirva a **cualquier obra**, no solo a la I.E.I. 881.

---

## Rol

Actúa como desarrollador full-stack con experiencia en control de obras públicas en el Perú (administración directa, Gobierno Regional de Puno, Invierte.pe, SIAF). Vas a construir una aplicación web que reemplace y unifique tres programas en Excel/VBA que ya existen en este repositorio.

## Contexto (lee esto antes de escribir código)

Lee en este orden y úsalo como especificación:

1. `README.md` – visión general y flujo.
2. `docs/01_FLUJO_GENERAL.md` – ciclo diario y mensual, actores.
3. `docs/02_MODELO_DE_DATOS.md` – entidades y campos (base del esquema de la base de datos).
4. `docs/04_CALCULO_VALORIZACION.md` – fórmulas que la web debe reproducir exactamente.
5. `docs/03_CONTENIDO_INFORME_MENSUAL.md` – cuadros del informe mensual (4.1 a 4.19, F-12B).
6. `programas/*/README.md` y el código VBA de `programas/*/src/*.bas` – la lógica actual (importación de S10/Delphin, estados de partidas, kardex, partes diarios). Cuando el código y la documentación difieran, manda el código.
7. `datos/iei881/valorizacion_01_2026-10.json` – datos reales para probar.

## Objetivo

Una aplicación web **multi-obra** donde el residente, el almacenero y el supervisor registran el avance diario y la aplicación genera la **valorización mensual** y los cuadros del informe, con la misma estructura y los mismos resultados que los libros de Excel.

## Alcance por módulos (construye en este orden; cada fase debe quedar funcionando)

### Fase 1 – Obra, presupuesto y valorización (prioridad)
1. **Obras**: crear/editar obra con todos los campos de `obra`, documentos (resoluciones, actas, ampliaciones de plazo), responsables y costos indirectos (% del C.D.). Selector de obra activa.
2. **Importar presupuesto y A.C.U.** desde Excel (S10, Delphin, Kennedy u otro): detecta encabezados de columnas y de partida («Partida 01.01.01» o «01.01.01 DESCRIPCIÓN»), rendimiento, secciones MO/MAT/EQ/SC/subpartidas, como hace `ModImportar.bas` del control de metrados. Vista previa antes de guardar y reporte de lo que no se pudo leer.
3. **Importar una valorización existente** desde el JSON de `datos/` (y desde el `.xlsx` de valorización con la misma lógica de `herramientas/exportar_valorizacion_json.py`) para cargar acumulados anteriores.
4. **Cronograma valorizado** (4.13): importar de Excel o digitar por partida y mes.
5. **Valorización mensual**: periodos (N° 01, 02…), cálculo de 4.4 por partida (anterior / actual / acumulado / saldo en metrado, S/ y %), títulos sumados por jerarquía, 4.3 resumen por subpresupuesto e indirectos, 4.3.1 Formato 10 (E.T. + MM + PN − D), 4.5 comparado vs programado, Curva S, % de cumplimiento. Cerrar periodo (bloquea edición y pasa actual → anterior).
6. **Panel de la obra**: tarjetas con presupuesto, valorizado acumulado, % físico, % programado, % financiero, días transcurridos/plazo, y la Curva S.

### Fase 2 – Control diario de metrados e insumos
7. **Metrado diario** por partida y día (vista tipo planilla del mes, como la hoja RESUMEN) y **planilla de metrados** con detalle por elemento (N° veces, N° elementos, largo, ancho, alto, diámetro, peso → parcial).
8. **Ingresos** (O/C, O/S, guías, actas de préstamo) con factor de conversión a la unidad del expediente.
9. **Consumo según A.C.U.** = metrado × cantidad unitaria; comparación con ingresos, HH del tareo y horas de equipos; estados de cada partida y día: `VÁLIDO`, `REVISAR / POR CONFIRMAR`, `SUPERA RENDIMIENTO`, `SUPERA METRADO DEL E.T.`, `NO VALORIZAR` (umbral configurable, 20 % por defecto). Los colores deben ser los mismos que en la hoja AYUDA del programa.

### Fase 3 – Almacén y equipo mecánico
10. Catálogo de insumos, movimientos (ingreso/salida con vale o PECOSA, pedido por, partida), saldo después de cada movimiento y alerta de saldo negativo.
11. **Kardex** y **bincard** por insumo; cuadros **4.10** movimiento diario, **4.11** ingresos – egresos, **4.12** saldo valorizado.
12. Equipos (O/S, unidad H.M./H.E./día, cantidad contratada, condición) y **partes diarios** (horario mañana/tarde, horómetro, combustible, partida, trabajo); cuadro **4.15** control de equipo y saldo de la O/S.
13. Opción «generar salidas y partes desde el A.C.U. × metrado del día» (como `ImportarControl`).

### Fase 4 – Salidas e informe
14. Exportar cada cuadro a **Excel** (mismo orden de columnas y encabezados que las hojas originales: rótulo con entidad, proyecto, CUI, meta, residente, supervisor, fechas) y a **PDF** (A4 horizontal; bincard y partes en A5 o A4 con dos por cara), con logo y firmas por cargo.
15. **Índice del informe mensual** con SI/NO, observación y responsable por ítem; checklist de lo que falta.
16. F-12B y cuadros financieros (PIM, certificado, devengado por mes y específica de gasto) con carga manual o desde Excel del SIAF.

## Requisitos técnicos

- **Stack** (consistente con la app existente `05_OBRAS\03_NORMATIVAS\webapp` – SGO-IEI881): Node.js ≥ 18, Express, **SQLite** (better-sqlite3) con migraciones, **ExcelJS** para importar/exportar, Multer para subir archivos. Frontend: HTML + JavaScript modular (o React + Vite si lo justificas), sin depender de servicios externos. Gráficos con Chart.js. PDF con Playwright/Chromium a partir de plantillas HTML de impresión.
- Debe correr **en local** en Windows con un `iniciar.bat` (`npm install` + `npm start`) y abrir en `http://localhost:3000`. Sin internet obligatorio.
- Todo en **español** (interfaz, mensajes, nombres de tablas y campos según `docs/02_MODELO_DE_DATOS.md`).
- Moneda S/ con 2 decimales y separador de miles; metrados con 2 decimales; porcentajes con 2 decimales. Los montos se redondean igual que en Excel (redondeo a 2 decimales por partida antes de sumar).
- Fechas `dd/mm/aaaa` en pantalla, ISO en la base. Calendario lunes a sábado, 8 h/día, feriados por obra.
- Roles simples: residente (todo), asistente técnico (metrados), almacenero (almacén y equipos), supervisor (solo lectura + conformidad), administrador.
- **Datos personales** (DNI, celular) opcionales y nunca en las exportaciones públicas.
- Respaldo: exportar/importar la obra completa en un solo JSON.
- Pruebas automáticas (Vitest o node:test) para todas las fórmulas de `docs/04_CALCULO_VALORIZACION.md`.

## Criterios de aceptación (con los datos de `datos/iei881/`)

Al importar `valorizacion_01_2026-10.json` la aplicación debe mostrar exactamente:

- 653 partidas y 247 títulos; C.D. del presupuesto **S/ 1 063 525.99**; presupuesto total **S/ 1 731 971.91**.
- Valorización N° 01 (15/10/2026 – 31/10/2026): C.D. del mes **S/ 54 534.96 (5.13 %)**; gastos generales S/ 10 546.06; supervisión S/ 6 667.76; total del mes con indirectos **S/ 88 811.20**.
- Subpresupuesto 01 acumulado S/ 30 095.90 (23.51 %); subpresupuesto 02 S/ 24 439.06 (6.22 %); 03 a 08 en cero.
- Programado octubre S/ 60 213.30 → cumplimiento **90.57 %**.
- Partida 01.01.01.01.01: 130.53 m² × S/ 1.97 = S/ 257.14, ejecutada 100 %, saldo 0.
- Curva S con 5 meses (OCT 2026 a FEB 2027) y % programado acumulado 5.66 %, 28.97 %, 62.79 %, 95.73 %, 100 %.

Además: crear una obra nueva vacía, importar un presupuesto de S10 y valorizar un mes sin tocar código.

## Forma de trabajo

1. Antes de programar, entrega: esquema de la base de datos (SQL), lista de pantallas con su ruta, y lista de endpoints de la API. Espera mi visto bueno.
2. Luego implementa **por fases**; al terminar cada fase: comandos para correrla, capturas o descripción de las pantallas, pruebas que pasan y qué queda pendiente.
3. Estructura propuesta del código, dentro de este repositorio:
   ```
   web/
     server.js
     src/db/ (migraciones, esquema)
     src/importadores/ (s10, delphin, valorizacion-json, valorizacion-xlsx)
     src/calculo/ (valorizacion, acu, almacen, equipos)
     src/exportadores/ (excel, pdf)
     src/rutas/
     public/ (pantallas)
     tests/
     iniciar.bat
   ```
4. No borres ni modifiques `programas/`, `datos/` ni `docs/`; si algo de la documentación está mal, propón el cambio aparte.
5. Si una regla no está clara, revisa el `.bas` correspondiente y cita el procedimiento (por ejemplo `ModActualizar.Calcular`) en tu explicación.
