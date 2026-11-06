#!/usr/bin/env python3
"""Lee los Excel de contadores desde OneDrive, los revisa y genera los datos del tablero.

Uso:
  python importar.py                 descarga de OneDrive y genera docs/datos.local.js
  python importar.py --carpeta RUTA  usa archivos ya descargados
  DATOS_CLAVE=... python importar.py además genera docs/datos.enc.json (cifrado, publicable)

Los Excel nunca se modifican. Las cifras son las que tienen las hojas diarias.
"""
import argparse
import base64
import calendar
import collections
import datetime as dt
import hashlib
import json
import os
import pathlib
import re
import secrets
import sys
import unicodedata

import openpyxl

BASE = pathlib.Path(__file__).resolve().parent
ANIOS = [int(a) for a in os.environ.get("ANIOS", "2026").split(",")]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
TOL = 0.5
RETENCION_MINIMA = 10
RETENCION_BAJA_SEGUIDA = 15
CAIDA_MINIMA = 0.40
DIAS_MINIMOS_MES_EN_CURSO = 3
DIA_LIMITE_MES_ANTERIOR = 10   # un recurrente pagado hasta este día puede corresponder al mes anterior
DIA_LIMITE_NOMINA = 7          # la nómina pagada hasta este día es de la quincena del mes anterior
MESES_PARA_INACTIVO = 3        # sin pagos en este tiempo, un recurrente deja de esperarse
ITERACIONES_CLAVE = 600_000

# Palabras de cada categoría. Gana la primera que coincida, por eso el orden importa.
CATEGORIAS = [
    ("Reposición de base", ["base"]),
    ("Vales y préstamos", ["vale", "prestamo", "preatamo", "adelanto"]),
    ("Nómina", ["nomin", "nomia", "nimina", "nomona", "prima ", "vacacion", "cesantia", "liquidacion", "turno",
                "turmo", "ayudante", "seguridad social", "liqud", "examenes"]),
    ("Arriendo", ["arrend", "arriend", "arremd", "pago local"]),
    ("Gaseosas", ["gaseosa", "postobon", "coca"]),
    ("Licor y cigarrillos", ["licor", "licar", " ron ", "aguardiente", "whisky", "vino", "cigarr", "cerveza", "tequila"]),
    ("Café y desechables", ["cafe", "azucar", "manantial", "desechable", "decechable", "aromatica", "agua", "vaso",
                            "mantial", "manantal", "platico", "plastico", "pitillo", "mezclador", "leche", "galleta", "hielo", "pinateria"]),
    ("Bonos y eventos", ["bono", "bingo", "rifa", "musico", "urna", "boleter", "evento", "torta", "decoracion",
                         "regalo", "cantante", "asado", "ecoracion", "ensalada", "arepa"]),
    ("Vigilancia", ["vigilancia", "vijilancia"]),
    ("Servicios e internet", ["tigo", "epm", "velonet", "internet", "recarga", "cargas ", "conexion", "coneccion",
                              "consignacion", "zona limpia", "enlinea", "en linea", "pago mes"]),
    ("Reparaciones", ["reparacion", "arreglo", "tarjeta", "repuesto", "tecnico", "cpu", "billetero", "llave",
                      "ferreteria", "mantenimiento", "bombillo", "cargador", "pintura", "electric", "plomer", "soldadura", "fuente", "fuemte", "monitor",
                      "targeta", "ventilador", "cable", "botoneria", "ensayo", "chapa", "extencion", "revision", "regrabado",
                      "progresivo", "riel", "amarradera", "candado", "canddo"]),
    ("Aseo", ["aseo", "limpion", "bolsa", "clorox", "alcohol", "ambientador", "basura", "jabon", "papel", "trapero", "escoba", "fumigacion", "limpiador", "veneno"]),
    ("Impuestos y honorarios", ["honorario", "asesoria", "dian", "sayco", "impuesto", "fee", "anexo", "coljuegos", "contador", "retencion", "camara y comercio", "programa sigo"]),
    ("Alimentación", ["almuerzo", "comida", "desayuno", "refrigerio", "domicilio"]),
]
NO_GASTO = {"Reposición de base"}


def num(x):
    return x if isinstance(x, (int, float)) and not isinstance(x, bool) else 0


def norm(x):
    s = unicodedata.normalize("NFKD", str(x or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def cop(x):
    return ("-" if round(x) < 0 else "") + "$" + f"{abs(x):,.0f}".replace(",", ".")


def categoria(concepto):
    n = " " + norm(concepto) + " "
    for nombre, palabras in CATEGORIAS:
        if any(p in n for p in palabras):
            return nombre
    return "Otros"


# ---------------------------------------------------------------- archivos
def periodo_de(nombre, carpeta=""):
    """Mes por el nombre del archivo (los títulos internos están desactualizados); año por nombre o carpeta."""
    n = norm(pathlib.Path(nombre).stem)
    mes = next((i for i, m in enumerate(MESES, 1) if m in n), None)
    if not mes:
        return None
    anio = re.search(r"(20\d\d)", n) or re.search(r"(20\d\d)", norm(carpeta))
    return (int(anio.group(1)), mes) if anio else None


def desde_onedrive(cache):
    sys.path.insert(0, str(BASE))
    from onedrive import OneDrive
    od = OneDrive()
    cache.mkdir(parents=True, exist_ok=True)
    candidatos = {}

    def considerar(item, carpeta):
        if not item["name"].lower().endswith((".xlsx", ".xlsm")):
            return
        p = periodo_de(item["name"], carpeta)
        if not p or p[0] not in ANIOS:
            return
        actual = candidatos.get(p)
        if not actual or item["lastModifiedDateTime"] > actual["lastModifiedDateTime"]:
            candidatos[p] = item

    def recorrer(id_, carpeta):
        for h in od.hijos(id_):
            if "folder" in h:
                recorrer(h["id"], carpeta)
            else:
                considerar(h, carpeta)

    for h in od.hijos():
        if "folder" in h and norm(h["name"]).startswith("contadores"):
            recorrer(h["id"], h["name"])
        elif "file" in h:
            considerar(h, "")

    archivos = []
    for periodo, item in sorted(candidatos.items()):
        destino = cache / f"{periodo[0]}-{periodo[1]:02d}.xlsx"
        marca = destino.with_suffix(".txt")
        if not (destino.exists() and marca.exists() and marca.read_text() == item["lastModifiedDateTime"]):
            od.descargar(item["id"], destino)
            marca.write_text(item["lastModifiedDateTime"])
        archivos.append(dict(ruta=destino, nombre=item["name"], periodo=periodo, modificado=item["lastModifiedDateTime"]))
    return archivos


def desde_carpeta(carpeta):
    archivos = []
    for p in sorted(pathlib.Path(carpeta).glob("*.xlsx")):
        periodo = periodo_de(p.name, p.parent.name)
        if periodo and periodo[0] in ANIOS:
            mod = dt.datetime.fromtimestamp(p.stat().st_mtime, dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            archivos.append(dict(ruta=p, nombre=p.name, periodo=periodo, modificado=mod))
    return sorted(archivos, key=lambda a: a["periodo"])


# ---------------------------------------------------------------- hojas
def columnas(ws):
    enc = {}
    for c in range(1, min(ws.max_column, 20) + 1):
        h = norm(ws.cell(2, c).value)
        if h and h not in enc:
            enc[h] = c
    if not all(k in enc for k in ("in ant", "total in", "out ant", "total out", "total yield", "total")):
        return None
    return dict(in_ant=enc["in ant"], in_act=enc["in ant"] + 1, tot_in=enc["total in"], out_ant=enc["out ant"],
                out_act=enc["out ant"] + 1, tot_out=enc["total out"], yield_=enc["total yield"], total=enc["total"],
                etiqueta=enc["total yield"] - 1, concepto=enc["total"] + 2, valor=enc["total"] + 3)


def dia_de_hoja(nombre):
    m = re.search(r"\d{1,2}", nombre)
    return int(m.group()) if m else None


def elegir_hojas(libro, anio, mes):
    """Asigna cada hoja a un día. Los archivos traen días del mes vecino y nombres errados, se usa el orden del libro."""
    ultimo = calendar.monthrange(anio, mes)[1]
    hojas = [(n, dia_de_hoja(n)) for n in libro.sheetnames if columnas(libro[n])]
    numeradas = [(n, d) for n, d in hojas if d]
    # Primera hoja del libro con el último día del mes anterior.
    while len(numeradas) > 1 and numeradas[0][1] >= 25 and numeradas[1][1] <= 3:
        numeradas.pop(0)
    elegidas, maximo = {}, 0
    for n, d in numeradas:
        if maximo >= 25 and d < 15:
            break  # desde aquí son días del mes siguiente
        if d <= ultimo:
            elegidas[d] = n  # si un día se repite gana la última versión del libro
            maximo = max(maximo, d)
    # Hoja sin número (ej. "vierne santo") que llena un hueco entre dos días.
    orden = [n for n, _ in hojas]
    for n, d in hojas:
        if d is None:
            i = orden.index(n)
            antes = next((dia_de_hoja(x) for x in reversed(orden[:i]) if dia_de_hoja(x)), None)
            despues = next((dia_de_hoja(x) for x in orden[i + 1:] if dia_de_hoja(x)), None)
            if antes and despues and despues - antes == 2 and antes + 1 not in elegidas:
                elegidas[antes + 1] = n
    return elegidas


def id_maquina(a):
    if isinstance(a, (int, float)) and not isinstance(a, bool):
        return str(int(a))
    s = str(a or "").strip().upper()
    return s if re.fullmatch(r"M\d{1,2}", s) else None


def id_sin_numero(filas, siguiente):
    """Fila con lecturas pero sin número: es la máquina que falta en la secuencia o un tramo de la anterior."""
    if not filas:
        return "1"
    anterior, sig = filas[-1]["m"], id_maquina(siguiente)
    if anterior.isdigit() and sig and sig.isdigit() and int(sig) - int(anterior) == 2:
        return str(int(anterior) + 1)
    return anterior


def nombre_valido(n):
    t = re.sub(r"\s+", " ", str(n or "")).strip()
    if not t or "borr" in norm(t):
        return ""
    return t.title()


def multiplicador(formula):
    m = re.search(r"\*\s*(\d+)", str(formula))
    return int(m.group(1)) if m else 1


def leer_mes(archivo):
    anio, mes = archivo["periodo"]
    v = openpyxl.load_workbook(archivo["ruta"], data_only=True)
    f = openpyxl.load_workbook(archivo["ruta"], data_only=False)
    elegidas = elegir_hojas(v, anio, mes)
    dias, sin_lectura = [], []
    for d in sorted(elegidas):
        hoja = elegidas[d]
        ws, wf = v[hoja], f[hoja]
        c = columnas(ws)
        filas, gastos, pie = [], [], {}
        for r in range(3, ws.max_row + 1):
            et = ws.cell(r, c["etiqueta"]).value
            if isinstance(et, str) and et.strip():
                pie.setdefault(norm(et), ws.cell(r, c["yield_"]).value)
            # La lista de gastos no siempre está en la misma columna: se busca a la derecha de TOTAL.
            for col in range(c["total"] + 1, c["total"] + 6):
                concepto, valor = ws.cell(r, col).value, ws.cell(r, col + 1).value
                if isinstance(concepto, str) and concepto.strip() and isinstance(valor, (int, float)) and valor:
                    gastos.append(dict(concepto=re.sub(r"\s+", " ", concepto).strip(), valor=valor))
                    break
            if not isinstance(ws.cell(r, c["in_ant"]).value, (int, float)):
                continue
            m = id_maquina(ws.cell(r, 1).value) or id_sin_numero(filas, ws.cell(r + 1, 1).value)
            crudo_in = wf.cell(r, c["tot_in"]).value
            k = multiplicador(crudo_in) if isinstance(crudo_in, str) else None
            filas.append(dict(
                m=m, nombre=nombre_valido(ws.cell(r, 2).value), k=k,
                leido=isinstance(ws.cell(r, c["in_act"]).value, (int, float)),
                in_ant=num(ws.cell(r, c["in_ant"]).value), in_act=num(ws.cell(r, c["in_act"]).value),
                out_ant=num(ws.cell(r, c["out_ant"]).value), out_act=num(ws.cell(r, c["out_act"]).value),
                entradas=num(ws.cell(r, c["tot_in"]).value), salidas=num(ws.cell(r, c["tot_out"]).value),
                resultado=num(ws.cell(r, c["yield_"]).value), total=num(ws.cell(r, c["total"]).value)))
        if not any(x["leido"] for x in filas):
            # Plantilla sin lecturas (ej. "22 no se trabajo"): sus gastos y efectivo son copia del día anterior.
            sin_lectura.append(dict(dia=d, hoja=hoja.strip()))
            continue
        dias.append(dict(dia=d, hoja=hoja.strip(), filas=filas, gastos=gastos, pie=pie))
    return dict(anio=anio, mes=mes, archivo=archivo["nombre"], modificado=archivo["modificado"], dias=dias, sin_lectura=sin_lectura)


# ---------------------------------------------------------------- análisis
def construir(meses):
    kmapa, nombres = collections.defaultdict(collections.Counter), collections.defaultdict(collections.Counter)
    for mes in meses:
        for d in mes["dias"]:
            for x in d["filas"]:
                if x["k"]:
                    kmapa[x["m"]][x["k"]] += 1
                if x["nombre"]:
                    nombres[x["m"]][x["nombre"]] += 1
    nombre = {m: c.most_common(1)[0][0] for m, c in nombres.items()}
    salida = {}
    previo = {}
    for mes in meses:
        salida[f"{mes['anio']}-{mes['mes']:02d}"] = analizar_mes(mes, kmapa, nombre, previo)
    return salida, nombre


def analizar_mes(mes, kmapa, nombre, previo):
    anio, nmes = mes["anio"], mes["mes"]
    alertas = []
    dias_out, gastos_out = [], []
    por_maquina = collections.defaultdict(lambda: dict(entradas=0, salidas=0, resultado=0, serie={}))
    saltos, difs_yield, gastos_descuadrados, pies_distintos = [], [], [], []

    def etiqueta(m):
        return f"Máquina {m}" + (f" ({nombre[m]})" if nombre.get(m) else "")

    for d in mes["dias"]:
        fecha = f"{anio}-{nmes:02d}-{d['dia']:02d}"
        entradas = salidas = resultado = 0
        actual = collections.defaultdict(list)
        for x in d["filas"]:
            m = x["m"]
            k = x["k"] or (kmapa[m].most_common(1)[0][0] if kmapa[m] else 1)
            entradas += x["entradas"]
            salidas += x["salidas"]
            resultado += x["resultado"]
            a = por_maquina[m]
            a["entradas"] += x["entradas"]
            a["salidas"] += x["salidas"]
            a["resultado"] += x["resultado"]
            s = a["serie"].setdefault(d["dia"], [0, 0, 0])
            s[0] += x["entradas"]
            s[1] += x["salidas"]
            s[2] += x["resultado"]
            actual[m].append(x)
            if abs(x["total"] - x["resultado"]) > TOL:
                difs_yield.append((fecha, m, x["total"] - x["resultado"]))
            if m in previo and (x["in_ant"] or x["out_ant"]):
                if not any(p["in_act"] == x["in_ant"] and p["out_act"] == x["out_ant"] for p in previo[m]):
                    p = previo[m][-1]
                    d_in, d_out = (x["in_ant"] - p["in_act"]) * k, (x["out_ant"] - p["out_act"]) * k
                    # Un valor menor al anterior es un borrado de la máquina, no plata sin contar.
                    saltos.append((fecha, m, max(d_in, 0), max(d_out, 0)))
        previo.clear()
        previo.update(actual)

        lista = d["gastos"]
        suma = sum(g["valor"] for g in lista)
        pie = d["pie"]
        if isinstance(pie.get("utilidad"), (int, float)) and abs(pie["utilidad"] - resultado) > 1:
            pies_distintos.append((fecha, pie["utilidad"], resultado))
        gastado = num(pie.get("gastos")) if "gastos" in pie else suma
        if abs(gastado - suma) > TOL:
            gastos_descuadrados.append((fecha, gastado, suma))
        for g in lista:
            gastos_out.append(dict(fecha=fecha, concepto=g["concepto"], categoria=categoria(g["concepto"]), valor=g["valor"]))
        efectivo = num(pie.get("efectivo"))
        esperado = num(pie.get("total")) if "total" in pie else resultado - gastado
        dias_out.append(dict(dia=d["dia"], hoja=d["hoja"], entradas=entradas, salidas=salidas, resultado=resultado,
                             gastos=gastado, esperado=esperado, efectivo=efectivo))

    # Días faltantes dentro del tramo ya digitado
    hechos = {d["dia"] for d in dias_out}
    tope = max(hechos) if hechos else 0
    sin_lectura = {x["dia"]: x["hoja"] for x in mes["sin_lectura"]}
    faltan = [d for d in range(1, tope + 1) if d not in hechos]
    for d in faltan:
        siguiente = [x for x in saltos if x[0] == f"{anio}-{nmes:02d}-{d + 1:02d}"]
        if siguiente:
            e, s_ = sum(x[2] for x in siguiente), sum(x[3] for x in siguiente)
            alertas.append(dict(nivel="revisar", titulo=f"Falta el día {d}: {cop(e - s_)} sin contar",
                                detalle=f"No hay hoja de ese día, pero los contadores sí avanzaron: entraron {cop(e)} y salieron {cop(s_)} que no aparecen en ningún total."))
            saltos = [x for x in saltos if x not in siguiente]
        else:
            nota = f"La hoja \"{sin_lectura[d]}\" no tiene lecturas" if d in sin_lectura else "No hay hoja"
            alertas.append(dict(nivel="info", titulo=f"Día {d} sin movimiento",
                                detalle=f"{nota} y los contadores siguen iguales al día siguiente, así que parece un día cerrado."))

    # Caja
    tot_res = sum(d["resultado"] for d in dias_out)
    tot_gas = sum(d["gastos"] for d in dias_out)
    tot_efe = sum(d["efectivo"] for d in dias_out)
    pagos_base = sum(g["valor"] for g in gastos_out if g["categoria"] in NO_GASTO)
    esperado = tot_res - tot_gas
    brecha = esperado - tot_efe
    if abs(brecha) > 0.01 * max(abs(esperado), 1):
        alertas.append(dict(nivel="revisar", titulo=f"{'Falta' if brecha > 0 else 'Sobra'} efectivo: {cop(abs(brecha))}",
                            detalle=f"Según las hojas debían quedar {cop(esperado)} y se contaron {cop(tot_efe)}."))

    if difs_yield:
        total = sum(x[2] for x in difs_yield)
        muestra = "; ".join(f"{x[0][8:]} máq. {x[1]} ({cop(x[2])})" for x in sorted(difs_yield, key=lambda z: -abs(z[2]))[:6])
        alertas.append(dict(nivel="revisar" if abs(total) >= 100_000 else "info",
                            titulo=f"YIELD distinto a entradas menos salidas en {len(difs_yield)} lecturas",
                            detalle=f"Suman {cop(total)}. Las mayores: {muestra}."))
    if saltos:
        e, s_ = sum(x[2] for x in saltos), sum(x[3] for x in saltos)
        fechas = sorted({x[0][8:] for x in saltos})
        alertas.append(dict(nivel="revisar" if e >= 100_000 else "info",
                            titulo=f"{len(saltos)} lecturas no continúan las del día anterior",
                            detalle=f"Días {', '.join(fechas)}. Entre un día y otro el contador avanzó sin quedar en ningún total: {cop(e)} de entradas y {cop(s_)} de salidas. Las máquinas borradas no suman aquí. " +
                                    "; ".join(f"máq. {x[1]} ({cop(x[2] - x[3])})" for x in sorted(saltos, key=lambda z: -abs(z[2] - z[3]))[:6]) + "."))
    for fecha, en_pie, en_filas in pies_distintos:
        alertas.append(dict(nivel="revisar" if abs(en_pie - en_filas) >= 100_000 else "info",
                            titulo=f"El total del día {fecha[8:]} no coincide con sus máquinas",
                            detalle=f"El pie de la hoja dice {cop(en_pie)} y la suma de las máquinas da {cop(en_filas)}. El tablero usa la suma de las máquinas."))
    if gastos_descuadrados:
        alertas.append(dict(nivel="info", titulo=f"En {len(gastos_descuadrados)} días el total de gastos no coincide con la lista",
                            detalle="; ".join(f"{x[0][8:]}: pie {cop(x[1])}, lista {cop(x[2])}" for x in gastos_descuadrados[:6]) + "."))

    maquinas = []
    for m, a in por_maquina.items():
        maquinas.append(dict(id=m, entradas=a["entradas"], salidas=a["salidas"], resultado=a["resultado"],
                             dias=[[d, *v] for d, v in sorted(a["serie"].items())]))
    maquinas.sort(key=lambda x: -x["resultado"])
    alertas.sort(key=lambda a: 0 if a["nivel"] == "revisar" else 1)
    return dict(anio=anio, mes=nmes, archivo=mes["archivo"], modificado=mes["modificado"], dias=dias_out,
                gastos=gastos_out, maquinas=maquinas, alertas=alertas,
                resumen=dict(entradas=sum(d["entradas"] for d in dias_out), salidas=sum(d["salidas"] for d in dias_out),
                             resultado=tot_res, gastos=tot_gas - pagos_base, base=pagos_base, pagos_caja=tot_gas, efectivo=tot_efe,
                             esperado=esperado, diferencia=brecha, sin_contar=sum(x[2] - x[3] for x in saltos),
                             utilidad_bruta=tot_res, utilidad_real=tot_res - (tot_gas - pagos_base), dias=len(dias_out),
                             dias_mes=calendar.monthrange(anio, nmes)[1]))


def item_recurrente(g):
    """Nombre del gasto recurrente al que pertenece un pago. Las reglas no llevan nombres de personas."""
    n, cat, dia = " " + norm(g["concepto"]) + " ", g["categoria"], int(g["fecha"][8:])
    if cat == "Arriendo":
        return "Arriendo bodega" if "bodega" in n or " apt" in n or (g["valor"] <= 1_000_000 and dia >= 20) else "Arriendo local"
    if cat == "Vigilancia":
        return "Vigilancia"
    if cat == "Gaseosas":
        return "Gaseosas"
    if cat == "Nómina":
        return "Nómina"
    if "honorario" in n or "asesoria" in n:
        return "Honorarios"
    if "tigo" in n:
        return "Tigo"
    if "velonet" in n:
        return "Velonet"
    if any(x in n for x in ("recarga", "cargas ", "conexion", "coneccion", "consignacion", "pago mes")):
        return "Conexión y recargas"
    return None


# El arriendo local se paga hacia el día 10 de forma regular: se deja en el mes de pago hasta confirmar a qué mes corresponde.
SIN_TRASLADO = {"Arriendo local", "Nómina"}
ORDEN_RECURRENTES = ["Nómina", "Arriendo local", "Arriendo bodega", "Vigilancia", "Honorarios", "Conexión y recargas", "Velonet", "Tigo", "Gaseosas"]


def es_quincena(g):
    n = norm(g["concepto"])
    return any(x in n for x in ("nomin", "nomia", "nimina", "nomona", "liqud")) and "vacacion" not in n


def asignar_gastos(meses):
    """Pasa al mes anterior los pagos que le corresponden. La caja sigue en la fecha real del pago."""
    claves = sorted(meses)
    posicion = {k: i for i, k in enumerate(claves)}

    def mover(g, origen, destino):
        meses[origen]["gastos"].remove(g)
        g["pagado_en"] = origen
        meses[destino]["gastos"].append(g)

    for m in meses.values():
        for g in m["gastos"]:
            item = item_recurrente(g)
            if item:
                g["item"] = item

    # Nómina de los primeros días: quincena del mes anterior
    for k in claves[1:]:
        for g in list(meses[k]["gastos"]):
            if g.get("item") == "Nómina" and int(g["fecha"][8:]) <= DIA_LIMITE_NOMINA and es_quincena(g):
                mover(g, k, claves[posicion[k] - 1])

    # Recurrentes: si al mes le falta el pago, lo toma de los primeros días del mes siguiente
    for item in ORDEN_RECURRENTES:
        if item in SIN_TRASLADO:
            continue
        for i, k in enumerate(claves[:-1]):
            ya_empezo = any(g.get("item") == item for c in claves[:i] for g in meses[c]["gastos"])
            if not ya_empezo or any(g.get("item") == item for g in meses[k]["gastos"]):
                continue
            siguiente = claves[i + 1]
            temprano = sorted((g for g in meses[siguiente]["gastos"] if g.get("item") == item and "pagado_en" not in g
                               and int(g["fecha"][8:]) <= DIA_LIMITE_MES_ANTERIOR), key=lambda g: g["fecha"])
            if temprano:
                mover(temprano[0], siguiente, k)

    # Totales del mes, ya con los gastos en su mes
    for k in claves:
        m = meses[k]
        m["gastos"].sort(key=lambda g: g["fecha"])
        entran = sum(g["valor"] for g in m["gastos"] if g.get("pagado_en") and g["categoria"] not in NO_GASTO)
        salen = sum(g["valor"] for c in claves if c != k for g in meses[c]["gastos"] if g.get("pagado_en") == k and g["categoria"] not in NO_GASTO)
        r = m["resumen"]
        r["gastos"] += entran - salen
        r["utilidad_real"] = r["utilidad_bruta"] - r["gastos"]
        r["gastos_de_otro_mes"] = entran
        r["gastos_a_otro_mes"] = salen

    # Tabla ordenada de recurrentes
    tabla = []
    for item in ORDEN_RECURRENTES:
        celdas = {}
        con_pago = [k for k in claves if any(g.get("item") == item for g in meses[k]["gastos"])]
        for i, k in enumerate(claves):
            pagos = [g for g in meses[k]["gastos"] if g.get("item") == item]
            completo = meses[k]["dias"] and meses[k]["dias"][-1]["dia"] == meses[k]["resumen"]["dias_mes"]
            if pagos:
                estado = "traido" if all(g.get("pagado_en") for g in pagos) else "pagado"
            elif not con_pago or k < con_pago[0]:
                estado = "antes"
            elif not completo:
                estado = "pendiente"
            elif k > con_pago[-1] and sum(1 for c in claves[posicion[con_pago[-1]] + 1:]
                                          if meses[c]["dias"][-1]["dia"] == meses[c]["resumen"]["dias_mes"]) > MESES_PARA_INACTIVO:
                estado = "inactivo"
            elif k > con_pago[-1] and con_pago[-1] < claves[-1] and i == len(claves) - 2:
                estado = "pendiente"  # se puede pagar en los primeros días del mes en curso
            else:
                estado = "falta"
            traidos = [g for g in pagos if g.get("pagado_en")]
            celdas[k] = dict(estado=estado, valor=sum(g["valor"] for g in pagos), traido=sum(g["valor"] for g in traidos))
        tabla.append(dict(item=item, celdas=celdas))
    return tabla


def vigilancia(meses, nombre):
    """Por cada mes, las máquinas que no retienen o que vienen mal, con los motivos en palabras."""
    claves = sorted(meses)

    def ret(x):
        return x["resultado"] / x["entradas"] * 100 if x["entradas"] else None

    def completo(m):
        return bool(m["dias"]) and m["dias"][-1]["dia"] == m["resumen"]["dias_mes"]

    def nom_mes(k):
        return MESES[int(k[5:]) - 1]

    for i, k in enumerate(claves):
        m = meses[k]
        casa = m["resumen"]["resultado"] / m["resumen"]["entradas"] * 100 if m["resumen"]["entradas"] else 0
        cerrados = [c for c in claves[:i + 1] if completo(meses[c])]
        lista = []
        for x in m["maquinas"]:
            motivos = []
            r = ret(x)
            if completo(m):
                if r is not None and r < RETENCION_MINIMA:
                    motivos.append(f"Retuvo {r:.0f} % en {nom_mes(k)}, la casa {casa:.0f} %")
            elif m["resumen"]["dias"] >= DIAS_MINIMOS_MES_EN_CURSO and x["resultado"] < 0:
                motivos.append(f"Va en pérdida este mes: {cop(x['resultado'])}")

            serie = []
            for c in cerrados:
                y = next((q for q in meses[c]["maquinas"] if q["id"] == x["id"]), None)
                if y:
                    serie.append((c, y))

            negativos = 0
            for c, y in reversed(serie):
                if y["resultado"] >= 0:
                    break
                negativos += 1
            if negativos >= 2:
                desde = serie[-negativos][0]
                total = sum(y["resultado"] for _, y in serie[-negativos:])
                motivos.append(f"En pérdida {negativos} meses seguidos desde {nom_mes(desde)}: {cop(total)}")

            bajos = 0
            for c, y in reversed(serie):
                v = ret(y)
                if v is None or v >= RETENCION_BAJA_SEGUIDA:
                    break
                bajos += 1
            if bajos >= 2 and negativos < 2:
                motivos.append(f"Retiene menos de {RETENCION_BAJA_SEGUIDA} % hace {bajos} meses")

            for campo, texto, formato in (("resultado", "Lo que deja", cop), ("entradas", "La venta", cop)):
                u = [y[campo] for _, y in serie[-4:]]
                if len(u) == 4 and u[0] > 0 and u[0] > u[1] > u[2] > u[3] and (u[0] - u[3]) / u[0] >= CAIDA_MINIMA:
                    motivos.append(f"{texto} cae 3 meses seguidos: de {formato(u[0])} en {nom_mes(serie[-4][0])} a {formato(u[3])} en {nom_mes(serie[-1][0])}")
            if motivos:
                lista.append(dict(id=x["id"], motivos=motivos, resultado=x["resultado"], retencion=r))
        lista.sort(key=lambda a: (-len(a["motivos"]), a["resultado"]))
        m["vigilar"] = lista


# ---------------------------------------------------------------- salida
def cifrar(texto, clave):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    sal, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    llave = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=sal, iterations=ITERACIONES_CLAVE).derive(clave.encode("utf-8"))
    datos = AESGCM(llave).encrypt(iv, texto.encode("utf-8"), None)
    b = lambda x: base64.b64encode(x).decode()
    return dict(v=1, iteraciones=ITERACIONES_CLAVE, sal=b(sal), iv=b(iv), datos=b(datos))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--carpeta", help="Lee de una carpeta en vez de OneDrive")
    args = ap.parse_args()
    archivos = desde_carpeta(args.carpeta) if args.carpeta else desde_onedrive(BASE / ".cache" / "excels")
    if not archivos:
        sys.exit("No se encontraron archivos de contadores para " + ", ".join(map(str, ANIOS)))

    leidos, errores = [], []
    for a in archivos:
        try:
            mes = leer_mes(a)
            if not mes["dias"]:
                raise ValueError("no encontré hojas diarias con las columnas esperadas")
            leidos.append(mes)
        except Exception as e:  # un archivo dañado no debe detener los demás meses
            errores.append(dict(archivo=a["nombre"], motivo=str(e)[:200]))

    meses, nombres = construir(leidos)
    recurrentes = asignar_gastos(meses)
    vigilancia(meses, nombres)
    # El registro de GitHub Actions es público: el detalle con cifras solo se imprime en el equipo local.
    for clave, m in ({} if os.environ.get("GITHUB_ACTIONS") else meses).items():
        r = m["resumen"]
        print(f"{clave}  {m['archivo']:<34} días {r['dias']:>2}  resultado {cop(r['resultado']):>15}  gastos {cop(r['gastos']):>14}  real {cop(r['utilidad_real']):>15}  alertas {len(m['alertas'])}  a vigilar {len(m['vigilar'])}")
    for e in errores:
        print("ERROR", e)

    datos = dict(generado=dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"), negocio="El Tesoro",
                 meses=meses, nombres=nombres, recurrentes=recurrentes, errores=errores)
    texto = json.dumps(datos, ensure_ascii=False, separators=(",", ":"))
    docs = BASE / "docs"
    docs.mkdir(exist_ok=True)
    (docs / "datos.local.js").write_text("window.DATOS=" + texto + ";\n", encoding="utf-8")
    clave = os.environ.get("DATOS_CLAVE")
    if clave:
        sin_fecha = json.dumps({k: v for k, v in datos.items() if k != "generado"}, ensure_ascii=False, sort_keys=True)
        huella = hashlib.sha256(sin_fecha.encode("utf-8")).hexdigest()
        marca = BASE / "estado" / "huella.txt"
        marca.parent.mkdir(exist_ok=True)
        if not (marca.exists() and marca.read_text().strip() == huella and (docs / "datos.enc.json").exists()):
            (docs / "datos.enc.json").write_text(json.dumps(cifrar(texto, clave)), encoding="utf-8")
            marca.write_text(huella)
            print("Datos cifrados en docs/datos.enc.json")
        else:
            print("Sin cambios en los datos")
    print(f"{len(meses)} meses, {len(errores)} con error, {len(texto) // 1024} KB")


if __name__ == "__main__":
    main()
