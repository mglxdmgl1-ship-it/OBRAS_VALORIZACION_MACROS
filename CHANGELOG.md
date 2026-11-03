# Registro de cambios

## 2026-11-02 – Organización del repositorio
- Se reúnen los programas de `05_OBRAS\04_INFORMES\10_PROGRAMAS_MACROS` en `programas/` (control de metrados, almacén y equipo, valorización mensual).
- Código VBA extraído de cada `.xlsm` a `src/` (`herramientas/extraer_vba.py`). El de almacén v008 coincide con la carpeta original `VBA_FUENTES`.
- Historial del programa de almacén v002 → v008 como commits de git (sus `.xlsm` antiguos quedan en la carpeta original).
- Valorización N° 01 (oct-2026) de la I.E.I. 881 exportada a JSON en `datos/iei881/`.
- Documentación: flujo, modelo de datos, contenido del informe, cálculo de la valorización e inventario del expediente técnico.
- Prompt para la página web en `prompts/PROMPT_PAGINA_WEB.md`.

## Versiones de los programas
| Programa | Versión vigente |
|---|---|
| Control de metrados diarios e insumos | Plantilla general v001 (obra 881: v003, mismo código) |
| Almacén y equipo mecánico | v008 |
| Valorización mensual I.E.I. 881 | Octubre 2026 v3 |
