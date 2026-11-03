"""
Exporta el libro de VALORIZACIÓN MENSUAL (formato GORE Puno, obra por administración directa)
a un JSON ordenado que la futura página web puede leer directamente.

Uso:
    pip install openpyxl
    python herramientas/exportar_valorizacion_json.py "programas/03_valorizacion_mensual/plantillas/VALORIZACION_MENSUAL_IEI881_2026-10_v3.xlsx" datos/iei881/valorizacion_01_2026-10.json

Hojas que lee (por nombre):
    BD                  datos generales de la obra, plazos, responsables, presupuesto
    4.4 Val.Pres P.B.   valorización por partida (anterior / actual / acumulado / saldo)
    4.3. ResVal P.B.    resumen por subpresupuesto y costos indirectos
    4.13 Cron.Val.      cronograma valorizado mensual por partida
    Curva-S             programado vs ejecutado (físico) vs devengado (financiero)

No exporta DNI ni números de celular de los responsables.
"""
import sys, json, datetime, re, warnings
import openpyxl

warnings.filterwarnings("ignore")


def val(v):
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    if isinstance(v, float):
        return round(v, 6)
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def celda(ws, ref):
    return val(ws[ref].value)


def buscar_fila(ws, col, texto, desde=1, hasta=None):
    hasta = hasta or ws.max_row
    for r in range(desde, hasta + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and texto.upper() in v.upper():
            return r
    return None


def leer_bd(ws):
    """Hoja BD: rótulos en la columna C, valores desde la columna D."""
    def fila(texto, col=3):
        return buscar_fila(ws, col, texto)

    def v(texto, col_val=4, col=3):
        r = fila(texto, col)
        return val(ws.cell(r, col_val).value) if r else None

    obra = {
        "nombre": v("OBRA:"),
        "nombre_corto": v("NOMBRE CORTO"),
        "meta": v("META (N°)"),
        "cui": v("CUI"),
        "unidad_ejecutora": v("UNIDAD EJECUTORA"),
        "pliego_ue_siaf": v("PLIEGO"),
        "area_responsable": v("ÁREA RESPONSABLE"),
        "modalidad": v("MODALIDAD DE EJECUCIÓN:"),
        "fuente_financiamiento": v("FUENTE DE FINANCIAMIENTO"),
        "funcion": v("FUNCIÓN:"),
        "division_funcional": v("DIVISIÓN FUNCIONAL"),
        "grupo_funcional": v("GRUPO FUNCIONAL"),
        "ubicacion": {
            "region": v("REGIÓN", 4),
            "provincia": v("PROVINCIA", 4),
            "distrito": v("DISTRITO", 4),
            "localidad": v("LOCALIDAD", 4),
            "coordenadas": v("COORDENADAS", 4),
            "altitud": v("ALTITUD", 4),
        },
    }
    # documentos de aprobación (C: documento, D: resolución, E: fecha, F: monto)
    docs = []
    r0 = fila("EXPEDIENTE TÉCNICO")
    if r0:
        for r in range(r0, r0 + 6):
            nombre = val(ws.cell(r, 3).value)
            if nombre:
                docs.append({"documento": nombre, "resolucion": val(ws.cell(r, 4).value),
                             "fecha": val(ws.cell(r, 5).value), "monto": val(ws.cell(r, 6).value)})
    # plazo
    plazo = {}
    r = fila("PLAZO DE EJECUCIÓN (ACTA")
    if r:
        plazo = {"documento": val(ws.cell(r, 4).value), "fecha": val(ws.cell(r, 5).value),
                 "dias": val(ws.cell(r, 6).value), "inicio": val(ws.cell(r, 7).value),
                 "termino": val(ws.cell(r, 8).value), "ampliaciones": []}
        for rr in range(r + 1, r + 7):
            n = val(ws.cell(rr, 3).value)
            if n and "AMPLIACIÓN" in str(n).upper():
                plazo["ampliaciones"].append({"nombre": n, "documento": val(ws.cell(rr, 4).value),
                                              "dias": val(ws.cell(rr, 6).value)})
            elif n and "TOTAL" in str(n).upper():
                plazo["plazo_total"] = val(ws.cell(rr, 6).value)
                plazo["termino_reprogramado"] = val(ws.cell(rr, 8).value)
    # responsables (sin DNI ni celular)
    responsables = []
    r = fila("RESIDENTE DE OBRA:")
    if r:
        for rr in range(r, r + 4):
            cargo = val(ws.cell(rr, 3).value)
            if cargo and str(cargo).endswith(":"):
                responsables.append({"cargo": cargo.rstrip(":"), "nombre": val(ws.cell(rr, 4).value),
                                     "profesion": val(ws.cell(rr, 5).value), "cip": val(ws.cell(rr, 6).value),
                                     "designacion": val(ws.cell(rr, 7).value)})
    # informe / periodo
    r = fila("INFORME MENSUAL N°")
    informe = {}
    if r:
        informe = {"numero": val(ws.cell(r, 4).value), "mes": val(ws.cell(r, 5).value), "anio": val(ws.cell(r, 6).value)}
        rp = fila("PERIODO DE LA VALORIZACIÓN")
        if rp:
            informe["periodo_del"] = val(ws.cell(rp, 4).value)
            informe["periodo_al"] = val(ws.cell(rp, 6).value)
        informe["valorizacion_n"] = v("VALORIZACIÓN N°")
        informe["mes_de_ejecucion"] = v("MES DE EJECUCIÓN")
    # costos del presupuesto (C: concepto, D: monto, E: % del C.D.)
    costos = []
    r = fila("COSTO DIRECTO")
    if r:
        for rr in range(r, r + 12):
            c = val(ws.cell(rr, 3).value)
            m = val(ws.cell(rr, 4).value)
            if c and isinstance(m, (int, float)):
                costos.append({"concepto": c, "monto": m, "pct_cd": val(ws.cell(rr, 5).value)})
    # feriados (K: fecha, L: descripción)
    feriados = []
    for rr in range(4, ws.max_row + 1):
        f = ws.cell(rr, 11).value
        if isinstance(f, datetime.datetime):
            feriados.append({"fecha": val(f), "descripcion": val(ws.cell(rr, 12).value)})
    return {"obra": obra, "documentos": docs, "plazo": plazo, "responsables": responsables,
            "informe": informe, "presupuesto": costos, "feriados": feriados}


def leer_partidas(ws):
    """Hoja 4.4: A TIPO (T=título, P=partida), B PADRE, C ITEM, D DESCRIPCIÓN, E UND, F METRADO, G P.U., H PARCIAL,
    I-K anterior, L-N actual, O-Q acumulado, R-T saldo (metrado, S/, %), V observación."""
    partidas, indirectos = [], []
    for r in range(14, ws.max_row + 1):
        tipo = ws.cell(r, 1).value
        item = val(ws.cell(r, 3).value)
        if tipo in ("T", "P"):
            def bloque(c):
                return {"metrado": val(ws.cell(r, c).value), "monto": val(ws.cell(r, c + 1).value),
                        "pct": val(ws.cell(r, c + 2).value)}
            partidas.append({
                "tipo": "titulo" if tipo == "T" else "partida",
                "item": item, "padre": None if ws.cell(r, 2).value in (None, "-") else val(ws.cell(r, 2).value),
                "nivel": len(str(item).split(".")),
                "descripcion": val(ws.cell(r, 4).value), "und": val(ws.cell(r, 5).value),
                "metrado": val(ws.cell(r, 6).value), "pu": val(ws.cell(r, 7).value), "parcial": val(ws.cell(r, 8).value),
                "anterior": bloque(9), "actual": bloque(12), "acumulado": bloque(15), "saldo": bloque(18),
                "observacion": val(ws.cell(r, 22).value),
            })
        elif item and isinstance(ws.cell(r, 8).value, (int, float)):
            indirectos.append({"concepto": item, "presupuesto": val(ws.cell(r, 8).value),
                               "anterior": val(ws.cell(r, 10).value), "actual": val(ws.cell(r, 13).value),
                               "acumulado": val(ws.cell(r, 16).value), "saldo": val(ws.cell(r, 19).value)})
    return partidas, indirectos


def leer_resumen(ws):
    """Hoja 4.3: B ITEM, C DESCRIPCIÓN, D MONTO, E-F anterior, G-H actual, I-J acumulado, K-L saldo.
    Filas con ITEM de 2 dígitos = subpresupuestos; filas con texto en B = totales (C.D., G.G., ... TOTAL)."""
    filas = []
    for r in range(14, ws.max_row + 1):
        b, c = val(ws.cell(r, 2).value), val(ws.cell(r, 3).value)
        monto = ws.cell(r, 4).value
        if b and isinstance(monto, (int, float)):
            es_sub = bool(re.match(r"^\d{2}$", str(b)))
            filas.append({"tipo": "subpresupuesto" if es_sub else "total",
                          "item": b if es_sub else None, "descripcion": c if es_sub else b,
                          "presupuesto": val(monto),
                          "anterior": val(ws.cell(r, 5).value), "actual": val(ws.cell(r, 7).value),
                          "pct_actual": val(ws.cell(r, 8).value),
                          "acumulado": val(ws.cell(r, 9).value), "pct_acumulado": val(ws.cell(r, 10).value),
                          "saldo": val(ws.cell(r, 11).value)})
    return filas


def leer_cronograma(ws):
    """Hoja 4.13: fila 11 = cabecera con los meses en G..K (o hasta la columna TOTAL)."""
    meses = []
    c = 7
    while ws.cell(11, c).value and not str(ws.cell(11, c).value).upper().startswith("TOTAL"):
        meses.append(str(ws.cell(11, c).value).split("\n")[0].strip())
        c += 1
    filas = []
    for r in range(12, ws.max_row + 1):
        item = val(ws.cell(r, 1).value)
        if not item or not re.match(r"^\d{2}(\.\d{2})*$", str(item)):
            continue
        filas.append({"item": item, "parcial": val(ws.cell(r, 6).value),
                      "por_mes": [val(ws.cell(r, 7 + i).value) or 0 for i in range(len(meses))]})
    return {"meses": meses, "partidas": filas}


def leer_curva_s(ws):
    filas = []
    for r in range(15, ws.max_row + 1):
        mes = val(ws.cell(r, 3).value)
        if isinstance(ws.cell(r, 2).value, (int, float)) and mes:
            filas.append({"n": ws.cell(r, 2).value, "mes": mes,
                          "programado": val(ws.cell(r, 4).value), "programado_pct_acum": val(ws.cell(r, 6).value),
                          "ejecutado": val(ws.cell(r, 7).value), "ejecutado_pct_acum": val(ws.cell(r, 9).value),
                          "devengado": val(ws.cell(r, 10).value), "devengado_pct_acum": val(ws.cell(r, 12).value)})
    return filas


def main(src, dst):
    wb = openpyxl.load_workbook(src, data_only=True)
    data = leer_bd(wb["BD"])
    partidas, indirectos = leer_partidas(wb["4.4 Val.Pres P.B."])
    data["valorizacion"] = {
        "partidas": partidas,
        "costos_indirectos": indirectos,
        "resumen_subpresupuestos": leer_resumen(wb["4.3. ResVal P.B."]),
    }
    data["cronograma_valorizado"] = leer_cronograma(wb["4.13 Cron.Val."])
    data["curva_s"] = leer_curva_s(wb["Curva-S"])
    data["_fuente"] = {"archivo": src.replace("\\", "/").split("/")[-1],
                       "exportado": datetime.datetime.now().isoformat(timespec="seconds"),
                       "nota": "Valores calculados del libro (data_only). Sin DNI ni celulares."}
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    n_p = sum(1 for p in partidas if p["tipo"] == "partida")
    print(f"OK {dst}: {n_p} partidas, {len(partidas) - n_p} títulos, "
          f"{len(data['cronograma_valorizado']['meses'])} meses, {len(data['curva_s'])} filas de curva S")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(1)
    main(sys.argv[1], sys.argv[2])
