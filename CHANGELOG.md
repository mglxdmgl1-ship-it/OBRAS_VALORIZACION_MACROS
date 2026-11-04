# Registro de cambios

## 2026-11-04 – Inicio de sesión y Supabase
- Página web con inicio de sesión (Supabase Auth) y dos roles: **administrador** (registra y edita todo, gestiona usuarios y accesos) y **cliente** (consulta sus obras y descarga reportes, solo lectura).
- Obras guardadas en Supabase (`sgo_obras`) con control de versión para cambios simultáneos, copia en el navegador para trabajar sin conexión y respaldo automático diario (`sgo_respaldos`).
- `supabase/migrations/`: esquema `sgo_*` con seguridad por fila; `supabase/functions/sgo-usuarios`: alta de cuentas, contraseñas temporales y retiro de accesos.
- Pantallas nuevas: ingreso, cambio de contraseña, Usuarios y accesos, Mi cuenta, Mis obras (cliente).

## 2026-11-03 – Página web
- `web/`: aplicación web estática (Netlify) que reúne la lógica de los tres programas: presupuesto, A.C.U., catálogo de precios, calculadora de insumos con redondeo y cantidades editables, metrados diarios con estados del control, almacén (salidas desde A.C.U. × metrado, kardex, bincard, 4.10–4.12), equipo mecánico (partes, 4.15), valorización mensual (4.3, 4.4, 4.5, curva S, cierre), cronograma 4.13 y reportes PDF / Excel con fórmulas.
- Importa presupuestos y A.C.U. de S10/Delphin (y de los propios libros de macros), metrados, movimientos, precios y cronogramas desde Excel, y el JSON de la valorización.
- `herramientas/exportar_obra_web.py`: arma el JSON de la obra de ejemplo para la web.
- Pruebas automáticas (`node --test web/tests/*.test.mjs`) con los criterios de aceptación del prompt.

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
