#!/usr/bin/env python3

import glob
import html
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

SHEET_NAME = "NACIONAL"          # hoja a procesar (si no existe, se usa la primera)
LISTA_SHEET = "Hoja1"            # hoja oculta con la lista desplegable de estados
OUTPUT_SUFFIX = "FIX"            # <nombre> + FIX + .xlsx

# Estado segun los miles del codigo de agencia (1010 -> 1, 15000 -> 15).
# Ortografia identica a Hoja1!B2:B25 (lista desplegable de la columna L).
ESTADOS = {
    1: "Distrito Capital", 2: "Amazonas", 3: "Anzoátegui", 4: "Apure",
    5: "Aragua", 6: "Barinas", 7: "Bolívar", 8: "Carabobo", 9: "Cojedes",
    10: "Delta Amacuro", 11: "Falcón", 12: "Guarico", 13: "Lara",
    14: "Merida", 15: "Miranda", 16: "Monagas", 17: "Nueva Esparta",
    18: "Portuguesa", 19: "Sucre", 20: "Tachira", 21: "Trujillo",
    22: "La Guaira", 23: "Yaracuy", 24: "Zulia",
}
# Codigos erroneos conocidos -> codigo correcto
CORRECCION_CODIGO = {22001: 22000}

# Columnas por defecto si no se encuentran los encabezados
COL_AGENCIA, COL_ESTADO, COL_TELEFONO, FILA_ENCABEZADO = "J", "L", "G", 3

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


# ----------------------------------------------------------------- utilidades
def col_to_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def sheet_paths(z):
    """Devuelve {nombre_hoja: ruta_xml} leyendo workbook.xml y sus relaciones."""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels.findall("pr:Relationship", NS)}
    out = {}
    for s in wb.find("m:sheets", NS):
        rid = s.get("{%s}id" % NS["r"])
        t = targets[rid].lstrip("/")
        out[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
    return out


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter("{%s}t" % NS["m"]))
            for si in root.findall("m:si", NS)]


CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
ROW_RE = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)


def attr(attrs, name):
    m = re.search(r'\b%s="([^"]*)"' % name, attrs)
    return m.group(1) if m else None


def cell_info(attrs, inner, sst):
    """(texto, es_numerico) de una celda."""
    inner = inner or ""
    t = attr(attrs, "t")
    if t == "inlineStr":
        txt = "".join(html.unescape(x) for x in re.findall(r"<t\b[^>]*>(.*?)</t>", inner, re.S))
        return txt, False
    v = re.search(r"<v>(.*?)</v>", inner, re.S)
    if not v:
        return "", False
    val = html.unescape(v.group(1))
    if t == "s":
        return sst[int(val)], False
    if t in ("str", "b", "e"):
        return val, False
    try:  # numerico
        f = float(val)
        if f == int(f):
            return str(int(f)), True
    except ValueError:
        pass
    return val, True


def make_cell(ref, style, text):
    s = ' s="%s"' % style if style else ""
    return '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (ref, s, escape(text))


# ------------------------------------------------------------------ reglas
def normaliza_agencia(valor):
    """-> (codigo_int | None). Acepta '01400', 1400, ' 1400 '."""
    digits = re.sub(r"\D", "", valor)
    return int(digits) if digits else None


def estado_de(codigo):
    return ESTADOS.get(codigo // 1000) if 1000 <= codigo <= 24999 else None


def varios_numeros(raw):
    """True si la celda parece traer mas de un telefono."""
    if re.search(r"[\n\r/,;]", raw):
        return True
    return sum(len(t) >= 7 for t in re.findall(r"\d+", raw)) >= 2


def formatea_telefono(raw):
    """Devuelve (resultado, estado); estado = 'ok' | 'invalido' | 'multiple'.
    Los validos no se alteran; los 'multiple' se dejan intactos."""
    if varios_numeros(raw):
        return raw, "multiple"
    d = re.sub(r"\D", "", raw)
    if len(d) == 11 and d[0] == "0":                       # Venezuela correcto
        return d, "ok"
    if len(d) == 10 and d[0] in "24":                      # falta el 0 inicial
        return "0" + d, "ok"
    if len(d) == 12 and d.startswith("58") and d[2] in "24":  # +58 ...
        return "0" + d[2:], "ok"
    if len(d) == 11 and d.startswith("51") and d[2] == "9":   # Peru con 51
        return d, "ok"
    if len(d) == 9 and d[0] == "9":                        # Peru sin 51
        return "51" + d, "ok"
    return "E" + d.zfill(10), "invalido"


# --------------------------------------------------------------- proceso
def lista_desplegable(z, paths, sst):
    """Valores de Hoja1!B2:B25 (None si la hoja no existe)."""
    if LISTA_SHEET not in paths:
        return None
    xml = z.read(paths[LISTA_SHEET]).decode("utf-8")
    vals = []
    for cm in CELL_RE.finditer(xml):
        ref = attr(cm.group(1), "r")
        if ref and re.match(r"B([2-9]|1\d|2[0-5])$", ref):
            vals.append(cell_info(cm.group(1), cm.group(2), sst)[0])
    return vals


def encuentra_columnas(xml, sst):
    cols = {}
    fila_enc = FILA_ENCABEZADO
    for rm in ROW_RE.finditer(xml):
        n = int(attr(rm.group(1), "r"))
        if n > 10:
            break
        for cm in CELL_RE.finditer(rm.group(2) or ""):
            txt = cell_info(cm.group(1), cm.group(2), sst)[0].strip().upper()
            col = split_ref(attr(cm.group(1), "r"))[0]
            if txt == "AGENCIA MRW":
                cols["J"] = col
            elif txt == "ESTADO DE DESTINO":
                cols["L"] = col
                fila_enc = n
            elif txt == "TELEFONO":
                cols["G"] = col
    return (cols.get("J", COL_AGENCIA), cols.get("L", COL_ESTADO),
            cols.get("G", COL_TELEFONO), fila_enc)


def procesa_hoja(xml, sst, rep):
    cJ, cL, cG, fila_enc = encuentra_columnas(xml, sst)
    cB = "B"  # NUM GUIA, solo para el reporte

    def fila(rm):
        attrs, inner = rm.group(1), rm.group(2)
        n = int(attr(attrs, "r"))
        if n <= fila_enc or inner is None:
            return rm.group(0)

        cells = []  # [idx, col, xml, attrs, inner]
        for cm in CELL_RE.finditer(inner):
            col = split_ref(attr(cm.group(1), "r"))[0]
            cells.append([col_to_idx(col), col, cm.group(0), cm.group(1), cm.group(2)])
        por_col = {c[1]: c for c in cells}

        def valor(col):
            c = por_col.get(col)
            return cell_info(c[3], c[4], sst) if c else ("", False)

        def escribe(col, texto):
            c = por_col.get(col)
            style = attr(c[3], "s") if c else None
            nueva = make_cell("%s%d" % (col, n), style, texto)
            if c:
                c[2] = nueva
            else:
                reg = [col_to_idx(col), col, nueva, "", None]
                cells.append(reg)
                por_col[col] = reg

        guia = valor(cB)[0]
        cambios = False

        # --- agencia (J) y estado (L)
        txtJ, numJ = valor(cJ)
        if txtJ.strip():
            codigo = normaliza_agencia(txtJ)
            if codigo is not None:
                codigo = CORRECCION_CODIGO.get(codigo, codigo)
            estado = estado_de(codigo) if codigo is not None else None
            if estado is None:
                rep["codigo_invalido"].append((n, guia, txtJ))
            else:
                nuevoJ = str(codigo).zfill(5)
                esta_ok = txtJ == nuevoJ
                if not esta_ok:
                    escribe(cJ, nuevoJ)
                    rep["agencia_reescrita"] += 1
                    cambios = True
                antes = valor(cL)[0]
                if antes != estado:
                    escribe(cL, estado)
                    rep["estado_escrito"] += 1
                    if antes:
                        rep["estado_corregido"].append((n, antes, estado))
                    cambios = True

        # --- telefono (G)
        txtG, _ = valor(cG)
        if txtG.strip():
            nuevo, estado_tel = formatea_telefono(txtG)
            if nuevo != txtG:
                escribe(cG, nuevo)
                rep["telefono_reescrito"] += 1
                cambios = True
            if estado_tel == "invalido":
                rep["telefono_invalido"].append((n, guia, txtG, nuevo))
            elif estado_tel == "multiple":
                rep["telefono_multiple"].append((n, guia, txtG))

        if not cambios:
            return rm.group(0)
        cells.sort(key=lambda c: c[0])
        return "<row%s>%s</row>" % (attrs, "".join(c[2] for c in cells))

    return ROW_RE.sub(fila, xml)


def procesa_archivo(ruta):
    base, _ = os.path.splitext(ruta)
    salida = base + OUTPUT_SUFFIX + ".xlsx"
    rep = {"agencia_reescrita": 0, "estado_escrito": 0, "telefono_reescrito": 0,
           "estado_corregido": [], "codigo_invalido": [], "telefono_invalido": [],
           "telefono_multiple": []}

    with zipfile.ZipFile(ruta) as zin:
        paths = sheet_paths(zin)
        nombre = SHEET_NAME if SHEET_NAME in paths else next(iter(paths))
        sst = shared_strings(zin)

        lista = lista_desplegable(zin, paths, sst)
        if lista is not None:
            faltan = sorted(set(ESTADOS.values()) - set(lista))
            if faltan:
                print("  AVISO: estos estados no estan en la lista desplegable: %s" % faltan)

        objetivo = paths[nombre]
        with zipfile.ZipFile(salida, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == objetivo:
                    data = procesa_hoja(data.decode("utf-8"), sst, rep).encode("utf-8")
                zout.writestr(item, data)

    print("\n%s  ->  %s" % (os.path.basename(ruta), os.path.basename(salida)))
    print("  Hoja: %s" % nombre)
    print("  Estados escritos:        %d" % rep["estado_escrito"])
    print("  Codigos de agencia ajustados: %d" % rep["agencia_reescrita"])
    print("  Telefonos reescritos:    %d" % rep["telefono_reescrito"])
    if rep["estado_corregido"]:
        print("  Estados que ya tenian otro valor y se sobrescribieron: %d" % len(rep["estado_corregido"]))
        for n, a, b in rep["estado_corregido"]:
            print("    fila %d: %s -> %s" % (n, a, b))
    if rep["codigo_invalido"]:
        print("  Codigos de agencia sin estado (L no modificada): %d" % len(rep["codigo_invalido"]))
        for n, guia, v in rep["codigo_invalido"]:
            print("    fila %d (%s): J=%r" % (n, guia, v))
    if rep["telefono_invalido"]:
        print("  Telefonos marcados con E: %d" % len(rep["telefono_invalido"]))
        for n, guia, a, b in rep["telefono_invalido"]:
            print("    fila %d (%s): %s -> %s" % (n, guia, a.replace("\n", " "), b))
    if rep["telefono_multiple"]:
        print("  Celdas con varios telefonos (sin modificar): %d" % len(rep["telefono_multiple"]))
        for n, guia, a in rep["telefono_multiple"]:
            print("    fila %d (%s): %s" % (n, guia, a.replace("\n", " | ")))


def main():
    archivos = sys.argv[1:]
    if not archivos:
        archivos = [f for f in glob.glob("*.xlsx")
                    if not os.path.basename(f).startswith("~$")
                    and not f.endswith(OUTPUT_SUFFIX + ".xlsx")]
    if not archivos:
        print("No se encontraron archivos .xlsx. Uso: python rellenar_estados.py archivo.xlsx")
        return 1
    for a in archivos:
        procesa_archivo(a)
    return 0


if __name__ == "__main__":
    sys.exit(main())
