# Valorización de obras por administración directa – programas y datos

Programas en Excel/VBA para llevar el control diario de una obra pública ejecutada por **administración directa** (formato del Gobierno Regional de Puno) y armar la **valorización mensual** y los cuadros del informe mensual. El repositorio reúne los programas, su código fuente VBA, el libro de valorización de referencia y los datos exportados a JSON, como base para construir después una **página web general** (sirve para cualquier obra, no solo para la de ejemplo).

Obra de ejemplo cargada: **I.E.I. N° 881 Miraflores – Cabanilla** (Lampa, Puno), CUI 2649805, meta 0180, presupuesto S/ 1 731 971.91, plazo 120 d.c. (15/10/2026 – 11/02/2027).

---

## Flujo general

```
EXPEDIENTE TÉCNICO (S10 / Delphin / Kennedy)
  presupuesto · A.C.U. · insumos · cronograma valorizado
        │  importar
        ▼
┌───────────────────────────────────────────┐
│ 1. CONTROL DE METRADOS DIARIOS E INSUMOS  │  metrado del día por partida → planillas diarias,
│    (programas/01_control_metrados_insumos)│  consumo de insumos por A.C.U., HH, maquinaria,
└───────────────────────────────────────────┘  alertas «no valorizar»
        │  importar (ingresos, salidas por A.C.U. × metrado, partes de equipos)
        ▼
┌───────────────────────────────────────────┐
│ 2. ALMACÉN Y EQUIPO MECÁNICO              │  kardex, bincard, partes diarios, cuadros 4.10,
│    (programas/02_almacen_equipo_mecanico) │  4.11, 4.12 y 4.15 del informe, PDF
└───────────────────────────────────────────┘
        │  metrados acumulados + cuadros de almacén/equipo
        ▼
┌───────────────────────────────────────────┐
│ 3. VALORIZACIÓN MENSUAL / INFORME         │  4.1 metrados, 4.2 planillas, 4.3 resumen,
│    (programas/03_valorizacion_mensual)    │  4.4 valorización, 4.5 comparado, Curva S,
└───────────────────────────────────────────┘  4.13 cronograma, F-12B
        │  exportar_valorizacion_json.py
        ▼
   datos/<obra>/valorizacion_NN_AAAA-MM.json  ──►  PÁGINA WEB (web/, publicada en Netlify)
```

Más detalle en [docs/01_FLUJO_GENERAL.md](docs/01_FLUJO_GENERAL.md).

## Contenido

| Carpeta | Qué hay |
|---|---|
| [`programas/01_control_metrados_insumos/`](programas/01_control_metrados_insumos/) | Plantilla general `.xlsm` (vacía, para cualquier obra), la copia cargada con la obra 881 y el código VBA (`src/`). |
| [`programas/02_almacen_equipo_mecanico/`](programas/02_almacen_equipo_mecanico/) | Programa de almacén y equipo mecánico v008, guía de uso y código VBA (`src/`). El historial de git guarda las versiones v002 → v008. |
| [`programas/03_valorizacion_mensual/`](programas/03_valorizacion_mensual/) | Libro de la valorización mensual N° 01 (octubre 2026) de la obra 881 y un reporte financiero de ejemplo (otra obra). |
| [`datos/iei881/`](datos/iei881/) | Valorización N° 01 exportada a JSON (653 partidas, cronograma, curva S, resumen). |
| [`docs/`](docs/) | Flujo, descripción de cada programa, cálculo de la valorización, modelo de datos para la web e inventario del expediente técnico. |
| [`herramientas/`](herramientas/) | `extraer_vba.py` (saca el VBA de un `.xlsm` sin abrir Excel) y `exportar_valorizacion_json.py`. |
| [`web/`](web/) | **Página web** de costos, presupuestos, A.C.U., calculadora de insumos, metrados diarios, almacén, equipos y valorización, con reportes PDF y Excel con fórmulas. Ver [web/README.md](web/README.md). |
| [`supabase/`](supabase/) | Base de datos en la nube de la página web: migraciones SQL (tablas `sgo_*`, permisos por rol con RLS, respaldo diario) y la función `sgo-usuarios`. |
| [`prompts/PROMPT_PAGINA_WEB.md`](prompts/PROMPT_PAGINA_WEB.md) | Prompt listo para pedir la página web de valorizaciones. |

## Uso rápido

**Programas en Excel**: abra el `.xlsm`, habilite macros y siga la hoja **INICIO** / **AYUDA** de cada uno. Orden mensual: preparar obra → registrar cada día → actualizar/revisar → generar formatos y PDF → cerrar mes.

**Volver a cargar el código en un libro** (si se modifica un `.bas`): en Excel, `Alt+F11` → *Archivo › Importar archivo…* → elegir el `.bas` de `src/` (reemplazar el módulo del mismo nombre). Al clonar en Windows los `.bas` salen en CP1252 + CRLF, que es lo que espera el editor de VBA (ver `.gitattributes`).

**Sacar el VBA de un libro** (para actualizar `src/` después de editar en Excel):

```bash
python herramientas/extraer_vba.py programas/02_almacen_equipo_mecanico/PROGRAMA_ALMACEN_EQUIPO_v008.xlsm salida_vba/
```

**Exportar la valorización a JSON** (para la web):

```bash
pip install openpyxl
python herramientas/exportar_valorizacion_json.py "programas/03_valorizacion_mensual/plantillas/VALORIZACION_MENSUAL_IEI881_2026-10_v3.xlsx" datos/iei881/valorizacion_01_2026-10.json
```

## Convenciones

- Versiones en el nombre del archivo: `_v001`, `_v002`… La versión vigente es la única binaria en el repositorio; las anteriores quedan en el historial de git (código fuente).
- Archivos de una obra concreta van en una subcarpeta con el nombre corto de la obra (`obra_iei881/`, `datos/iei881/`).
- Moneda: soles (S/). Fechas en ISO (`AAAA-MM-DD`) en los JSON.

## Privacidad

Los libros de la obra 881 contienen nombres, CIP, DNI y celulares de los responsables, además de montos y órdenes de compra. **Mantenga este repositorio como privado.** El JSON exportado ya excluye DNI y celulares.

## Origen

Carpeta original: `05_OBRAS\04_INFORMES\10_PROGRAMAS_MACROS` (expediente técnico en `05_OBRAS\01_ET_INSTITUCION_EDUCATIVA`, que no se sube por su tamaño, ~3.8 GB; ver [docs/05_EXPEDIENTE_TECNICO.md](docs/05_EXPEDIENTE_TECNICO.md)).
