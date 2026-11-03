#!/usr/bin/env python3
"""Arma el JSON de una obra para la página web (web/public/datos/<obra>.json).

Une en un solo archivo:
  * la valorización exportada (datos/<obra>/valorizacion_NN_AAAA-MM.json): datos de la obra,
    presupuesto, costos indirectos, cronograma valorizado, curva S;
  * el control de metrados (.xlsm): PRESUPUESTO (rendimiento), ACU, RESUMEN (metrados de cada día)
    e INGRESOS (para enlazar los ítems del almacén con los insumos del A.C.U.);
  * el programa de almacén y equipo (.xlsm): CATALOGO, MOVIMIENTOS, EQUIPOS y PARTES.

Uso:
    python herramientas/exportar_obra_web.py VALORIZACION.json CONTROL.xlsm ALMACEN.xlsm salida.json

No exporta DNI, celulares ni N° de CIP (la web es pública).
"""
import datetime as dt
import json
import re
import sys

import openpyxl


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return 0.0


def fecha(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return (dt.date(1899, 12, 30) + dt.timedelta(days=int(v))).isoformat()
    return None


def norm(s):
    s = txt(s).upper()
    for a, b in zip("ÁÉÍÓÚÜ", "AEIOUU"):
        s = s.replace(a, b)
    return re.sub(r"[^A-Z0-9]", "", s)


MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def mes_iso(s):
    p = s.upper().replace("SEPTIEMBRE", "SETIEMBRE").split()
    return f"{p[1]}-{MESES.index(p[0]) + 1:02d}"


def main(f_val, f_ctrl, f_alm, f_out):
    val = json.load(open(f_val, encoding="utf-8"))
    ctrl = openpyxl.load_workbook(f_ctrl, data_only=True, read_only=True)
    alm = openpyxl.load_workbook(f_alm, data_only=True, read_only=True)

    o = val["obra"]
    u = o.get("ubicacion", {})
    plazo = val["plazo"]
    cd = next(p["monto"] for p in val["presupuesto"] if p["concepto"] == "COSTO DIRECTO")
    ind = []
    for p in val["presupuesto"]:
        c = p["concepto"]
        if c in ("COSTO DIRECTO", "TOTAL COSTO INDIRECTO", "PRESUPUESTO TOTAL") or c.startswith("PIM"):
            continue
        if isinstance(p.get("pct_cd"), (int, float)):
            ind.append({"concepto": c, "monto": p["monto"]})

    # ---------------- presupuesto (del control: trae el rendimiento) ----------------
    rend = {}
    ws = ctrl["PRESUPUESTO"]
    for r in ws.iter_rows(min_row=8, values_only=True):
        it = txt(r[0])
        if it:
            rend[it] = num(r[6]) if r[6] is not None else None

    partidas = []
    for p in val["valorizacion"]["partidas"]:
        es_p = p["tipo"] == "partida"
        partidas.append({
            "item": p["item"],
            "tipo": "P" if es_p else "T",
            "descripcion": txt(p["descripcion"]),
            "und": txt(p["und"]) if es_p else "",
            "metrado": p["metrado"] if es_p else None,
            "pu": p["pu"] if es_p else None,
            "rendimiento": rend.get(p["item"]) if es_p else None,
            "origen": "ET",
        })
    items_p = {p["item"] for p in partidas if p["tipo"] == "P"}

    # ---------------- A.C.U. ----------------
    acu = {}
    insumos = {}
    ws = ctrl["ACU"]
    for r in ws.iter_rows(min_row=8, values_only=True):
        it = txt(r[0])
        if not it:
            continue
        tipo = txt(r[3]).upper() or "MAT"
        nom = txt(r[1])
        fila = {"tipo": tipo, "cantidad": num(r[4])}
        if tipo == "%MO":
            fila["descripcion"] = nom
            fila["und"] = "%mo"
        else:
            k = nom.upper()
            if k not in insumos:
                insumos[k] = {"codigo": "", "descripcion": nom, "und": txt(r[2]), "tipo": tipo,
                              "precio": num(r[5]), "controlar": True}
            fila["insumo"] = insumos[k]["descripcion"]
        acu.setdefault(it, []).append(fila)

    # códigos del catálogo de insumos del A.C.U.: MO-, MT-, EQ-, SC-
    pref = {"MO": "MO", "MAT": "MT", "EQ": "EQ", "SC": "SC", "SP": "SP"}
    cont = {}
    no_alm = ["AGUA PUESTA EN OBRA"]
    for k in sorted(insumos, key=lambda x: (insumos[x]["tipo"], x)):
        i = insumos[k]
        pfx = pref.get(i["tipo"], "OT")
        cont[pfx] = cont.get(pfx, 0) + 1
        i["codigo"] = f"{pfx}-{cont[pfx]:04d}"
        if any(w in k for w in no_alm):
            i["controlar"] = False

    # ---------------- metrados diarios (RESUMEN, columnas 1…31 del mes que se controla) ----------------
    d = ctrl["DATOS"]
    mes = None
    for r in d.iter_rows(min_row=1, max_row=45, values_only=True):
        if r[1] and "MES QUE SE CONTROLA" in txt(r[1]).upper():
            mes = r[2]
    mes = mes.date() if isinstance(mes, dt.datetime) else dt.date(2026, 10, 1)
    metrados = {}
    ws = ctrl["RESUMEN"]
    for r in ws.iter_rows(min_row=10, values_only=True):
        if txt(r[0]).upper() != "P":
            continue
        it = txt(r[1])
        for dd in range(1, 32):
            v = num(r[7 + dd])
            if v > 0:
                try:
                    f = dt.date(mes.year, mes.month, dd).isoformat()
                except ValueError:
                    continue
                metrados.setdefault(f, {})[it] = v

    # ---------------- ingresos del control: enlace documento → insumo del A.C.U. ----------------
    enlace = {}
    ws = ctrl["INGRESOS"]
    for r in ws.iter_rows(min_row=8, values_only=True):
        desc, ins = txt(r[5]), txt(r[8])
        if desc and ins:
            fa = r[10]
            enlace[norm(desc)] = (ins, num(fa) if fa not in (None, "") else None)

    # ---------------- almacén: catálogo, movimientos ----------------
    almacen = []
    cods = {}
    ws = alm["CATALOGO"]
    for r in ws.iter_rows(min_row=8, values_only=True):
        cod = txt(r[1])
        if not cod:
            continue
        acu_nom = txt(r[14])
        fac = num(r[15]) if r[15] not in (None, "") else None
        e = enlace.get(norm(r[2]))
        if e:
            if not acu_nom:
                acu_nom = e[0]
            if fac is None and e[1]:
                fac = e[1]
        if acu_nom and acu_nom.upper() not in insumos:
            acu_nom = ""
        it = {"codigo": cod, "descripcion": txt(r[2]), "und": txt(r[3]), "grupo": txt(r[4]) or "MATERIALES",
              "pu": num(r[5]), "insumo_acu": acu_nom or "", "factor": fac or 1, "observacion": txt(r[16])}
        if acu_nom:
            it["insumo_acu"] = insumos[acu_nom.upper()]["descripcion"]
        almacen.append(it)
        cods[cod] = it

    movimientos = []
    ws = alm["MOVIMIENTOS"]
    n = 0
    for r in ws.iter_rows(min_row=8, values_only=True):
        f = fecha(r[1])
        tipo = txt(r[2]).upper()
        if not f or tipo not in ("INGRESO", "SALIDA", "DEVOLUCION", "DEVOLUCIÓN"):
            continue
        n += 1
        movimientos.append({
            "id": f"M{n:05d}", "fecha": f, "tipo": "DEVOLUCION" if tipo.startswith("DEVOL") else tipo,
            "documento": txt(r[3]), "numero": txt(r[4]), "oc_os": txt(r[5]), "procedencia": txt(r[6]),
            "partida": txt(r[7]), "codigo": txt(r[8]), "cantidad": num(r[11]), "pedido_por": txt(r[12]),
            "vale": "", "observacion": txt(r[13]), "estado": "CONFIRMADO", "origen": txt(r[14]) or "DIGITADO",
        })

    # ---------------- equipos y partes ----------------
    equipos = []
    ws = alm["EQUIPOS"]
    for r in ws.iter_rows(min_row=8, values_only=True):
        cod = txt(r[1])
        if not cod:
            continue
        os_txt = txt(r[8])
        m = re.search(r"(\d{2})/(\d{2})/(\d{4})", os_txt)
        os_f = f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None
        acu_nom = txt(r[23]) if len(r) > 23 else ""
        if acu_nom.upper() in insumos:
            acu_nom = insumos[acu_nom.upper()]["descripcion"]
        else:
            acu_nom = ""
        equipos.append({
            "codigo": cod, "descripcion": txt(r[2]), "marca": txt(r[3]), "modelo": txt(r[4]), "serie": txt(r[5]),
            "anio": txt(r[6]), "proveedor": txt(r[7]), "os": os_txt, "os_fecha": os_f, "unidad": txt(r[9]) or "H.M.",
            "cantidad": num(r[10]), "costo": num(r[11]), "condicion": txt(r[13]), "combustible": txt(r[14]),
            "operador": txt(r[15]), "insumo_acu": acu_nom,
        })
    # fechas de inicio de la O/S según el control (INGRESOS)
    ws = ctrl["INGRESOS"]
    for r in ws.iter_rows(min_row=8, values_only=True):
        ins, f = txt(r[8]), fecha(r[1])
        for e in equipos:
            if e["insumo_acu"] and ins.upper() == e["insumo_acu"].upper() and f:
                e["inicio"] = f
                e["fin_os"] = fecha(r[13])
                e["estado"] = txt(r[14]) or "CONFIRMADO"

    partes = []
    ws = alm["PARTES"]
    n = 0
    for r in ws.iter_rows(min_row=8, values_only=True):
        f = fecha(r[1])
        if not f or not txt(r[2]):
            continue
        n += 1
        partes.append({
            "id": f"P{n:05d}", "fecha": f, "equipo": txt(r[2]), "n_parte": txt(r[4]), "operador": txt(r[5]),
            "m_de": txt(r[6]), "m_a": txt(r[7]), "t_de": txt(r[8]), "t_a": txt(r[9]),
            "h_ini": r[10] if r[10] is not None else None, "h_fin": r[11] if r[11] is not None else None,
            "partida": txt(r[12]), "trabajo": txt(r[13]), "metrado": num(r[14]),
            "horas": num(r[20]) if len(r) > 20 and r[20] is not None else None,
            "combustible": None, "observacion": txt(r[21]) if len(r) > 21 else "",
        })

    # ---------------- cronograma y curva S ----------------
    cr = val["cronograma_valorizado"]
    meses = [mes_iso(m) for m in cr["meses"]]
    cron = {}
    for p in cr["partidas"]:
        if p["item"] in items_p:
            cron[p["item"]] = [round(x or 0, 2) for x in p["por_mes"]]
    total = next(p["monto"] for p in val["presupuesto"] if p["concepto"] == "PRESUPUESTO TOTAL")
    devengado = {}
    for i, c in enumerate(val["curva_s"]):
        if c.get("devengado_pct_acum") is not None and i < len(meses):
            devengado[meses[i]] = round(c["devengado_pct_acum"] * total, 2)

    inf = val["informe"]
    obra = {
        "version": 1,
        "id": "iei881",
        "datos": {
            "nombre": txt(o["nombre"]).strip("“”\""),
            "nombre_corto": txt(o["nombre_corto"]),
            "cui": txt(o["cui"]), "meta": txt(o["meta"]),
            "entidad": txt(o.get("unidad_ejecutora")) or "GOBIERNO REGIONAL DE PUNO",
            "dependencia": txt(o.get("area_responsable")),
            "pliego": txt(o.get("pliego_ue_siaf")),
            "modalidad": txt(o.get("modalidad")) or "ADMINISTRACIÓN DIRECTA",
            "fuente_financiamiento": txt(o.get("fuente_financiamiento")),
            "funcion": txt(o.get("funcion")), "division_funcional": txt(o.get("division_funcional")),
            "grupo_funcional": txt(o.get("grupo_funcional")),
            "region": txt(u.get("region")), "provincia": txt(u.get("provincia")), "distrito": txt(u.get("distrito")),
            "localidad": txt(u.get("localidad")),
            "fecha_inicio": plazo["inicio"], "plazo_dias": plazo["dias"],
            "ampliacion_dias": sum(a.get("dias") or 0 for a in plazo.get("ampliaciones", [])),
            "presupuesto_total_resolucion": total,
        },
        "parametros": {"umbral": 0.2, "jornada": 8, "uso_max_hh": 0.85, "no_almacen": "AGUA PUESTA EN OBRA",
                       "combustibles": "DIESEL;PETROLEO;GASOHOL;GASOLINA;KEROSENE", "redondeo": "arriba",
                       "dias_laborables": [1, 2, 3, 4, 5, 6]},
        "responsables": [{"cargo": r["cargo"], "nombre": r["nombre"], "profesion": r.get("profesion") or ""}
                         for r in val["responsables"]],
        "firmas": ["RESIDENTE DE OBRA", "SUPERVISOR DE OBRA", ""],
        "indirectos": ind,
        "documentos": [{"documento": x["documento"], "resolucion": x.get("resolucion") or "",
                        "fecha": x.get("fecha"), "monto": x.get("monto")} for x in val["documentos"]],
        "feriados": val.get("feriados", []),
        "partidas": partidas,
        "acu": acu,
        "insumos": sorted(insumos.values(), key=lambda x: x["codigo"]),
        "metrados": metrados,
        "detalles": {},
        "almacen": almacen,
        "movimientos": movimientos,
        "equipos": equipos,
        "partes": partes,
        "periodos": [{"numero": int(inf["valorizacion_n"]), "del": inf["periodo_del"], "al": inf["periodo_al"],
                      "estado": "ABIERTO", "manual": {}, "cerrado": None}],
        "anterior_inicial": {},
        "cronograma": {"meses": meses, "montos": cron},
        "financiero": {"devengado_acumulado": devengado},
        "creado": dt.datetime.now().isoformat(timespec="seconds"),
    }
    json.dump(obra, open(f_out, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))

    # ---------------- comprobaciones ----------------
    s_cd = round(sum(round(p["metrado"] * p["pu"], 2) for p in partidas if p["tipo"] == "P"), 2)
    s_mes = 0.0
    tot_met = {}
    for f, m in metrados.items():
        if inf["periodo_del"] <= f <= inf["periodo_al"]:
            for it, v in m.items():
                tot_met[it] = tot_met.get(it, 0) + v
    pu = {p["item"]: p["pu"] for p in partidas if p["tipo"] == "P"}
    s_mes = round(sum(round(v * pu.get(it, 0), 2) for it, v in tot_met.items()), 2)
    ref = {p["item"]: p["actual"]["metrado"] for p in val["valorizacion"]["partidas"] if p["tipo"] == "partida"}
    dif = [(k, tot_met.get(k, 0), v) for k, v in ref.items() if abs((v or 0) - tot_met.get(k, 0)) > 0.005]
    print(f"partidas {len(items_p)}, títulos {len(partidas) - len(items_p)}, C.D. {s_cd:,.2f} (ref {cd:,.2f})")
    print(f"A.C.U.: {len(acu)} partidas, {sum(len(v) for v in acu.values())} filas, {len(insumos)} insumos")
    print(f"metrados diarios: {len(metrados)} días; valorizado del periodo {s_mes:,.2f}")
    print(f"diferencias de metrado del mes con la valorización: {len(dif)} {dif[:5]}")
    print(f"almacén: {len(almacen)} ítems ({sum(1 for a in almacen if a['insumo_acu'])} enlazados al A.C.U.), "
          f"{len(movimientos)} movimientos; equipos {len(equipos)}, partes {len(partes)}")


if __name__ == "__main__":
    if len(sys.argv) != 5:
        print(__doc__)
        sys.exit(1)
    main(*sys.argv[1:])
