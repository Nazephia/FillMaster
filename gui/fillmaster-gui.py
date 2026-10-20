#!/usr/bin/env python3

import html
import os
import queue
import re
import subprocess
import sys
import threading
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

APP_TITULO = "Relleno de estados MRW"
SHEET_NAME = "NACIONAL"
LISTA_SHEET = "Hoja1"

# Estado según los miles del código de agencia (1010 -> 1, 15000 -> 15).
# Ortografía idéntica a Hoja1!B2:B25 (lista desplegable de la columna L).
ESTADOS = {
    1: "Distrito Capital", 2: "Amazonas", 3: "Anzoátegui", 4: "Apure",
    5: "Aragua", 6: "Barinas", 7: "Bolívar", 8: "Carabobo", 9: "Cojedes",
    10: "Delta Amacuro", 11: "Falcón", 12: "Guarico", 13: "Lara",
    14: "Merida", 15: "Miranda", 16: "Monagas", 17: "Nueva Esparta",
    18: "Portuguesa", 19: "Sucre", 20: "Tachira", 21: "Trujillo",
    22: "La Guaira", 23: "Yaracuy", 24: "Zulia",
}
CORRECCION_CODIGO = {22001: 22000}   # códigos erróneos conocidos
COL_AGENCIA, COL_ESTADO, COL_TELEFONO, FILA_ENCABEZADO = "J", "L", "G", 3

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


# =====================================================================
#  NÚCLEO (misma lógica que rellenar_estados.py)
# =====================================================================
def col_to_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def sheet_paths(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels.findall("pr:Relationship", NS)}
    out = {}
    for s in wb.find("m:sheets", NS):
        t = targets[s.get("{%s}id" % NS["r"])].lstrip("/")
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
        return "".join(html.unescape(x) for x in re.findall(r"<t\b[^>]*>(.*?)</t>", inner, re.S)), False
    v = re.search(r"<v>(.*?)</v>", inner, re.S)
    if not v:
        return "", False
    val = html.unescape(v.group(1))
    if t == "s":
        return sst[int(val)], False
    if t in ("str", "b", "e"):
        return val, False
    try:
        f = float(val)
        if f == int(f):
            return str(int(f)), True
    except ValueError:
        pass
    return val, True


def make_cell(ref, style, text):
    s = ' s="%s"' % style if style else ""
    return '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (ref, s, escape(text))


def normaliza_agencia(valor):
    digits = re.sub(r"\D", "", valor)
    return int(digits) if digits else None


def estado_de(codigo):
    return ESTADOS.get(codigo // 1000) if 1000 <= codigo <= 24999 else None


def varios_numeros(raw):
    if re.search(r"[\n\r/,;]", raw):
        return True
    return sum(len(t) >= 7 for t in re.findall(r"\d+", raw)) >= 2


def formatea_telefono(raw):
    """-> (resultado, estado) con estado 'ok' | 'invalido' | 'multiple'."""
    if varios_numeros(raw):
        return raw, "multiple"
    d = re.sub(r"\D", "", raw)
    if len(d) == 11 and d[0] == "0":                          # Venezuela correcto
        return d, "ok"
    if len(d) == 10 and d[0] in "24":                         # falta el 0 inicial
        return "0" + d, "ok"
    if len(d) == 12 and d.startswith("58") and d[2] in "24":  # +58 ...
        return "0" + d[2:], "ok"
    if len(d) == 11 and d.startswith("51") and d[2] == "9":   # Perú con 51
        return d, "ok"
    if len(d) == 9 and d[0] == "9":                           # Perú sin 51
        return "51" + d, "ok"
    return "E" + d.zfill(10), "invalido"


def lista_desplegable(z, paths, sst):
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
    cols, fila_enc = {}, FILA_ENCABEZADO
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
    encontrados = all(k in cols for k in ("J", "L", "G"))
    return (cols.get("J", COL_AGENCIA), cols.get("L", COL_ESTADO),
            cols.get("G", COL_TELEFONO), fila_enc, encontrados)


def procesa_hoja(xml, sst, rep, opts):
    cJ, cL, cG, fila_enc, encontrados = encuentra_columnas(xml, sst)
    if not encontrados:
        rep["avisos"].append("No se encontraron todos los encabezados; se usaron las "
                             "columnas %s (agencia), %s (estado) y %s (teléfono)." % (cJ, cL, cG))

    def fila(rm):
        attrs, inner = rm.group(1), rm.group(2)
        n = int(attr(attrs, "r"))
        if n <= fila_enc or inner is None:
            return rm.group(0)

        cells = []
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

        guia = valor("B")[0]
        cambios = False

        txtJ, _ = valor(cJ)
        if txtJ.strip() and (opts["estado"] or opts["codigos"]):
            codigo = normaliza_agencia(txtJ)
            if codigo is not None:
                codigo = CORRECCION_CODIGO.get(codigo, codigo)
            estado = estado_de(codigo) if codigo is not None else None
            if estado is None:
                rep["codigo_invalido"].append((n, guia, txtJ))
            else:
                if opts["codigos"]:
                    nuevoJ = str(codigo).zfill(5)
                    if txtJ != nuevoJ:
                        escribe(cJ, nuevoJ)
                        rep["agencia_reescrita"] += 1
                        cambios = True
                if opts["estado"]:
                    antes = valor(cL)[0]
                    if antes != estado:
                        escribe(cL, estado)
                        rep["estado_escrito"] += 1
                        if antes:
                            rep["estado_corregido"].append((n, antes, estado))
                        cambios = True

        txtG, _ = valor(cG)
        if opts["telefonos"] and txtG.strip():
            nuevo, est = formatea_telefono(txtG)
            if nuevo != txtG:
                escribe(cG, nuevo)
                rep["telefono_reescrito"] += 1
                cambios = True
            if est == "invalido":
                rep["telefono_invalido"].append((n, guia, txtG, nuevo))
            elif est == "multiple":
                rep["telefono_multiple"].append((n, guia, txtG))

        if not cambios:
            return rm.group(0)
        cells.sort(key=lambda c: c[0])
        return "<row%s>%s</row>" % (attrs, "".join(c[2] for c in cells))

    return ROW_RE.sub(fila, xml)


def nuevo_reporte():
    return {"agencia_reescrita": 0, "estado_escrito": 0, "telefono_reescrito": 0,
            "estado_corregido": [], "codigo_invalido": [], "telefono_invalido": [],
            "telefono_multiple": [], "avisos": [], "hoja": ""}


def procesa_archivo(ruta, salida, opts):
    """Lee `ruta`, escribe `salida` y devuelve el reporte. No toca el original."""
    rep = nuevo_reporte()
    with zipfile.ZipFile(ruta) as zin:
        paths = sheet_paths(zin)
        if SHEET_NAME in paths:
            nombre = SHEET_NAME
        else:
            nombre = next(iter(paths))
            rep["avisos"].append("No existe la hoja «%s»; se procesó la primera hoja («%s»)."
                                 % (SHEET_NAME, nombre))
        rep["hoja"] = nombre
        sst = shared_strings(zin)

        lista = lista_desplegable(zin, paths, sst)
        if lista is not None:
            faltan = sorted(set(ESTADOS.values()) - set(lista))
            if faltan:
                rep["avisos"].append("Estos estados no están en la lista desplegable "
                                     "de la columna L: %s" % ", ".join(faltan))

        objetivo = paths[nombre]
        tmp = salida + ".tmp"
        try:
            with zipfile.ZipFile(tmp, "w") as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == objetivo:
                        data = procesa_hoja(data.decode("utf-8"), sst, rep, opts).encode("utf-8")
                    zout.writestr(item, data)
            os.replace(tmp, salida)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)
    return rep


# =====================================================================
#  INTERFAZ
# =====================================================================
TEMAS = {
    "claro": dict(
        fondo="#F3F3F3", panel="#FFFFFF", texto="#1F1F1F", suave="#605E5C",
        acento="#2B579A", acento_hover="#1E3F73", acento_texto="#FFFFFF",
        borde="#D0D0D0", campo="#FFFFFF", seleccion="#CFE0F5", seleccion_texto="#1F1F1F",
        deshab="#A6A6A6", boton="#E6E6E6", boton_hover="#D6D6D6",
        ok="#107C10", aviso="#B35A00", error="#C42B1C", cabecera="#2B579A"),
    "oscuro": dict(
        fondo="#1E1E1E", panel="#2B2B2B", texto="#EDEDED", suave="#A8A8A8",
        acento="#3A78C2", acento_hover="#5290DA", acento_texto="#FFFFFF",
        borde="#454545", campo="#202020", seleccion="#2F4B73", seleccion_texto="#FFFFFF",
        deshab="#6B6B6B", boton="#3A3A3A", boton_hover="#4A4A4A",
        ok="#6CCB5F", aviso="#FFB347", error="#FF7B72", cabecera="#1F3A63"),
}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITULO)
        self.geometry("880x800")
        self.minsize(780, 680)

        self.tema = "claro"
        self.cola = queue.Queue()
        self.archivos = []
        self.ultima_carpeta = None
        self.trabajando = False

        self.v_estado = tk.BooleanVar(value=True)
        self.v_codigos = tk.BooleanVar(value=True)
        self.v_telefonos = tk.BooleanVar(value=True)
        self.v_destino = tk.StringVar(value="original")
        self.v_carpeta = tk.StringVar(value="")
        self.v_sufijo = tk.StringVar(value="FIX")

        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=10)
        tkfont.nametofont("TkTextFont").configure(size=10)
        self.f_titulo = tkfont.Font(family=base.cget("family"), size=16, weight="bold")
        self.f_seccion = tkfont.Font(family=base.cget("family"), size=10, weight="bold")
        self.f_mono = tkfont.Font(family="Consolas" if sys.platform == "win32" else "DejaVu Sans Mono",
                                  size=9)

        self.estilo = ttk.Style(self)
        self.estilo.theme_use("clam")
        self._construir()
        self.aplicar_tema()
        self.after(100, self._revisar_cola)

    # ------------------------------------------------------------ UI
    def _tarjeta(self, padre, titulo):
        """Panel con borde y título; devuelve el marco interior."""
        ext = tk.Frame(padre, bd=0, highlightthickness=1)
        self.tarjetas.append(ext)
        ttk.Label(ext, text=titulo, style="Titulo.TLabel").pack(anchor="w", padx=12, pady=(8, 0))
        inte = ttk.Frame(ext, style="Tarjeta.TFrame", padding=(12, 6, 12, 12))
        inte.pack(fill="both", expand=True)
        return ext, inte

    def _construir(self):
        self.tarjetas = []
        # Cabecera
        self.cab = tk.Frame(self)
        self.cab.pack(fill="x")
        self.lbl_titulo = tk.Label(self.cab, text=APP_TITULO, font=self.f_titulo, anchor="w")
        self.lbl_titulo.pack(side="left", padx=18, pady=12)
        self.btn_tema = tk.Button(self.cab, text="Modo oscuro", relief="flat", bd=0,
                                  padx=12, pady=4, cursor="hand2", command=self.cambiar_tema)
        self.btn_tema.pack(side="right", padx=18)

        cuerpo = ttk.Frame(self, style="Fondo.TFrame", padding=(16, 12, 16, 12))
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=1)
        cuerpo.rowconfigure(3, weight=1)

        # 1. Archivos
        c1, s1 = self._tarjeta(cuerpo, "1. Archivos de Excel")
        c1.grid(row=0, column=0, sticky="ew")
        s1.columnconfigure(0, weight=1)
        marco_lista = ttk.Frame(s1, style="Tarjeta.TFrame")
        marco_lista.grid(row=0, column=0, sticky="ew")
        marco_lista.columnconfigure(0, weight=1)
        self.lista = tk.Listbox(marco_lista, height=5, selectmode="extended", activestyle="none",
                                relief="flat", highlightthickness=1, bd=0)
        self.lista.grid(row=0, column=0, sticky="ew")
        sb = ttk.Scrollbar(marco_lista, orient="vertical", command=self.lista.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.lista.configure(yscrollcommand=sb.set)
        botones = ttk.Frame(s1, style="Tarjeta.TFrame")
        botones.grid(row=0, column=1, sticky="n", padx=(10, 0))
        self.b_agregar = ttk.Button(botones, text="Agregar archivos…", style="Acento.TButton",
                                    command=self.agregar)
        self.b_agregar.pack(fill="x")
        self.b_quitar = ttk.Button(botones, text="Quitar seleccionados", command=self.quitar)
        self.b_quitar.pack(fill="x", pady=(6, 0))
        self.b_limpiar = ttk.Button(botones, text="Vaciar lista", command=self.limpiar)
        self.b_limpiar.pack(fill="x", pady=(6, 0))

        # 2. Opciones
        c2, s2 = self._tarjeta(cuerpo, "2. Opciones")
        c2.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        s2.columnconfigure(1, weight=1)
        ttk.Label(s2, text="¿Qué se debe hacer?", style="Seccion.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(s2, text="Rellenar «Estado de destino» según el código de agencia",
                        variable=self.v_estado, style="Tarjeta.TCheckbutton").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Checkbutton(s2, text="Ajustar códigos de agencia (cero delante en los de 4 dígitos)",
                        variable=self.v_codigos, style="Tarjeta.TCheckbutton").grid(
            row=2, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(s2, text="Dar formato a los teléfonos (Venezuela y Perú)",
                        variable=self.v_telefonos, style="Tarjeta.TCheckbutton").grid(
            row=3, column=0, columnspan=3, sticky="w")

        ttk.Label(s2, text="¿Dónde guardar los archivos nuevos?", style="Seccion.TLabel").grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Radiobutton(s2, text="Junto a cada archivo original", value="original",
                        variable=self.v_destino, style="Tarjeta.TRadiobutton",
                        command=self._estado_carpeta).grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Radiobutton(s2, text="En otra carpeta:", value="otra", variable=self.v_destino,
                        style="Tarjeta.TRadiobutton", command=self._estado_carpeta).grid(
            row=6, column=0, sticky="w")
        self.e_carpeta = ttk.Entry(s2, textvariable=self.v_carpeta, state="readonly")
        self.e_carpeta.grid(row=6, column=1, sticky="ew", padx=(8, 8))
        self.b_carpeta = ttk.Button(s2, text="Elegir carpeta…", command=self.elegir_carpeta)
        self.b_carpeta.grid(row=6, column=2)

        ttk.Label(s2, text="Texto que se agrega al nombre:", style="Tarjeta.TLabel").grid(
            row=7, column=0, sticky="w", pady=(10, 0))
        self.e_sufijo = ttk.Entry(s2, textvariable=self.v_sufijo, width=14)
        self.e_sufijo.grid(row=7, column=1, sticky="w", padx=(8, 0), pady=(10, 0))
        ttk.Label(s2, text="Ejemplo: MASTER.xlsx → MASTERFIX.xlsx. El original nunca se modifica.",
                  style="Suave.TLabel").grid(row=8, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self._estado_carpeta()

        # 3. Procesar
        s3 = ttk.Frame(cuerpo, style="Fondo.TFrame")
        s3.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        s3.columnconfigure(1, weight=1)
        self.b_procesar = ttk.Button(s3, text="Procesar archivos", style="Acento.TButton",
                                     command=self.procesar)
        self.b_procesar.grid(row=0, column=0, ipadx=16, ipady=4)
        self.progreso = ttk.Progressbar(s3, mode="determinate")
        self.progreso.grid(row=0, column=1, sticky="ew", padx=(14, 0))
        self.v_estado_txt = tk.StringVar(value="Agregue uno o más archivos para comenzar.")
        ttk.Label(s3, textvariable=self.v_estado_txt, style="Estado.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Resultado
        c4, s4 = self._tarjeta(cuerpo, "Resultado")
        c4.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        c4.rowconfigure(1, weight=1)
        s4.columnconfigure(0, weight=1)
        s4.rowconfigure(0, weight=1)
        self.texto = tk.Text(s4, wrap="word", relief="flat", highlightthickness=1, bd=0,
                             font=self.f_mono, height=10, state="disabled", padx=8, pady=6)
        self.texto.grid(row=0, column=0, sticky="nsew")
        sb2 = ttk.Scrollbar(s4, orient="vertical", command=self.texto.yview)
        sb2.grid(row=0, column=1, sticky="ns")
        self.texto.configure(yscrollcommand=sb2.set)
        pie = ttk.Frame(s4, style="Tarjeta.TFrame")
        pie.grid(row=1, column=0, columnspan=2, sticky="e", pady=(8, 0))
        self.b_guardar = ttk.Button(pie, text="Guardar reporte…", command=self.guardar_reporte)
        self.b_guardar.pack(side="left")
        self.b_abrir = ttk.Button(pie, text="Abrir carpeta de resultados", command=self.abrir_carpeta,
                                  state="disabled")
        self.b_abrir.pack(side="left", padx=(8, 0))

    # ------------------------------------------------------------ tema
    def cambiar_tema(self):
        self.tema = "oscuro" if self.tema == "claro" else "claro"
        self.aplicar_tema()

    def aplicar_tema(self):
        c = TEMAS[self.tema]
        st = self.estilo
        self.configure(bg=c["fondo"])

        st.configure(".", background=c["fondo"], foreground=c["texto"], fieldbackground=c["campo"],
                     bordercolor=c["borde"], lightcolor=c["borde"], darkcolor=c["borde"],
                     troughcolor=c["boton"], focuscolor=c["acento"])
        st.configure("Fondo.TFrame", background=c["fondo"])
        st.configure("Tarjeta.TFrame", background=c["panel"])
        st.configure("Titulo.TLabel", background=c["panel"], foreground=c["acento"], font=self.f_seccion)
        st.configure("Estado.TLabel", background=c["fondo"], foreground=c["suave"])
        for t in self.tarjetas:
            t.configure(bg=c["panel"], highlightbackground=c["borde"], highlightcolor=c["borde"])
        st.configure("Tarjeta.TLabel", background=c["panel"], foreground=c["texto"])
        st.configure("Seccion.TLabel", background=c["panel"], foreground=c["texto"], font=self.f_seccion)
        st.configure("Suave.TLabel", background=c["panel"], foreground=c["suave"])
        st.configure("TLabel", background=c["fondo"], foreground=c["texto"])
        # etiquetas "suaves" sobre el fondo (debajo del botón Procesar)
        st.configure("Tarjeta.TCheckbutton", background=c["panel"], foreground=c["texto"],
                     indicatorbackground=c["campo"], indicatorforeground=c["acento"])
        st.map("Tarjeta.TCheckbutton", background=[("active", c["panel"])],
               indicatorbackground=[("selected", c["acento"]), ("active", c["campo"])],
               indicatorforeground=[("selected", c["acento_texto"])],
               foreground=[("disabled", c["deshab"])])
        st.configure("Tarjeta.TRadiobutton", background=c["panel"], foreground=c["texto"],
                     indicatorbackground=c["campo"], indicatorforeground=c["acento_texto"])
        st.map("Tarjeta.TRadiobutton", background=[("active", c["panel"])],
               indicatorbackground=[("selected", c["acento"]), ("active", c["campo"])],
               foreground=[("disabled", c["deshab"])])

        st.configure("TButton", background=c["boton"], foreground=c["texto"], bordercolor=c["borde"],
                     padding=(10, 5), relief="flat")
        st.map("TButton", background=[("disabled", c["boton"]), ("pressed", c["boton_hover"]),
                                      ("active", c["boton_hover"])],
               foreground=[("disabled", c["deshab"])])
        st.configure("Acento.TButton", background=c["acento"], foreground=c["acento_texto"],
                     bordercolor=c["acento"], padding=(10, 5), font=self.f_seccion)
        st.map("Acento.TButton", background=[("disabled", c["borde"]), ("pressed", c["acento_hover"]),
                                             ("active", c["acento_hover"])],
               foreground=[("disabled", c["deshab"])])

        st.configure("TEntry", fieldbackground=c["campo"], foreground=c["texto"], insertcolor=c["texto"],
                     bordercolor=c["borde"], padding=4)
        st.map("TEntry", fieldbackground=[("readonly", c["campo"]), ("disabled", c["panel"])],
               foreground=[("disabled", c["deshab"]), ("readonly", c["texto"])])
        st.configure("Horizontal.TProgressbar", background=c["acento"], troughcolor=c["boton"],
                     bordercolor=c["boton"], lightcolor=c["acento"], darkcolor=c["acento"])
        st.configure("Vertical.TScrollbar", background=c["boton"], troughcolor=c["panel"],
                     bordercolor=c["panel"], arrowcolor=c["texto"])
        st.map("Vertical.TScrollbar", background=[("active", c["boton_hover"])])

        # Widgets clásicos de tk
        self.cab.configure(bg=c["cabecera"])
        self.lbl_titulo.configure(bg=c["cabecera"], fg="#FFFFFF")
        self.btn_tema.configure(bg=c["acento"] if self.tema == "claro" else c["acento"],
                                fg="#FFFFFF", activebackground=c["acento_hover"],
                                activeforeground="#FFFFFF",
                                text="Modo oscuro" if self.tema == "claro" else "Modo claro")
        self.lista.configure(bg=c["campo"], fg=c["texto"], selectbackground=c["seleccion"],
                             selectforeground=c["seleccion_texto"], highlightbackground=c["borde"],
                             highlightcolor=c["acento"])
        self.texto.configure(bg=c["campo"], fg=c["texto"], insertbackground=c["texto"],
                             selectbackground=c["seleccion"], selectforeground=c["seleccion_texto"],
                             highlightbackground=c["borde"], highlightcolor=c["acento"])
        self.texto.tag_configure("titulo", foreground=c["acento"],
                                 font=(self.f_mono.cget("family"), 10, "bold"))
        self.texto.tag_configure("ok", foreground=c["ok"])
        self.texto.tag_configure("aviso", foreground=c["aviso"])
        self.texto.tag_configure("error", foreground=c["error"])
        self.texto.tag_configure("suave", foreground=c["suave"])

    # ------------------------------------------------------------ acciones
    def _estado_carpeta(self):
        estado = "normal" if self.v_destino.get() == "otra" else "disabled"
        self.b_carpeta.configure(state=estado)

    def agregar(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccione los archivos de Excel",
            filetypes=[("Libros de Excel", "*.xlsx"), ("Todos los archivos", "*.*")])
        for r in rutas:
            r = os.path.abspath(r)
            if r not in self.archivos:
                self.archivos.append(r)
                self.lista.insert("end", r)
        self._actualiza_estado()

    def quitar(self):
        for i in reversed(self.lista.curselection()):
            self.lista.delete(i)
            del self.archivos[i]
        self._actualiza_estado()

    def limpiar(self):
        self.lista.delete(0, "end")
        self.archivos.clear()
        self._actualiza_estado()

    def _actualiza_estado(self):
        n = len(self.archivos)
        self.v_estado_txt.set("%d archivo(s) en la lista." % n if n
                              else "Agregue uno o más archivos para comenzar.")

    def elegir_carpeta(self):
        d = filedialog.askdirectory(title="Carpeta donde guardar los resultados")
        if d:
            self.v_carpeta.set(os.path.abspath(d))

    def abrir_carpeta(self):
        if not self.ultima_carpeta or not os.path.isdir(self.ultima_carpeta):
            return
        try:
            if sys.platform == "win32":
                os.startfile(self.ultima_carpeta)           # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self.ultima_carpeta])
            else:
                subprocess.Popen(["xdg-open", self.ultima_carpeta])
        except Exception as e:
            messagebox.showerror(APP_TITULO, "No se pudo abrir la carpeta:\n%s" % e)

    def guardar_reporte(self):
        contenido = self.texto.get("1.0", "end").strip()
        if not contenido:
            messagebox.showinfo(APP_TITULO, "Todavía no hay un reporte para guardar.")
            return
        ruta = filedialog.asksaveasfilename(
            title="Guardar reporte", defaultextension=".txt", initialfile="reporte.txt",
            filetypes=[("Texto", "*.txt")])
        if ruta:
            with open(ruta, "w", encoding="utf-8-sig") as f:
                f.write(contenido + "\n")

    # ------------------------------------------------------------ reporte
    def _escribe(self, texto, tag=None):
        self.texto.configure(state="normal")
        self.texto.insert("end", texto + "\n", tag)
        self.texto.see("end")
        self.texto.configure(state="disabled")

    def _bloquea(self, bloqueado):
        estado = "disabled" if bloqueado else "normal"
        for b in (self.b_procesar, self.b_agregar, self.b_quitar, self.b_limpiar):
            b.configure(state=estado)
        self.trabajando = bloqueado

    # ------------------------------------------------------------ proceso
    def _ruta_salida(self, ruta):
        base, _ = os.path.splitext(os.path.basename(ruta))
        carpeta = (self.v_carpeta.get() if self.v_destino.get() == "otra"
                   else os.path.dirname(ruta))
        return os.path.join(carpeta, base + self.v_sufijo.get().strip() + ".xlsx")

    def procesar(self):
        if self.trabajando:
            return
        if not self.archivos:
            messagebox.showinfo(APP_TITULO, "Primero agregue al menos un archivo de Excel.")
            return
        opts = {"estado": self.v_estado.get(), "codigos": self.v_codigos.get(),
                "telefonos": self.v_telefonos.get()}
        if not any(opts.values()):
            messagebox.showinfo(APP_TITULO, "Marque al menos una opción en «¿Qué se debe hacer?».")
            return
        if self.v_destino.get() == "otra" and not os.path.isdir(self.v_carpeta.get()):
            messagebox.showinfo(APP_TITULO, "Elija la carpeta donde se guardarán los resultados.")
            return

        trabajos = [(r, self._ruta_salida(r)) for r in self.archivos]
        if any(os.path.normcase(os.path.abspath(o)) == os.path.normcase(os.path.abspath(r))
               for r, o in trabajos):
            messagebox.showwarning(
                APP_TITULO,
                "El resultado tendría el mismo nombre que el original y lo reemplazaría.\n"
                "Escriba un texto en «Texto que se agrega al nombre» o elija otra carpeta.")
            return
        existentes = [o for _, o in trabajos if os.path.exists(o)]
        if existentes:
            if not messagebox.askyesno(
                    APP_TITULO, "Ya existen %d archivo(s) de resultado con ese nombre.\n"
                                "¿Desea reemplazarlos?" % len(existentes)):
                return

        self.texto.configure(state="normal")
        self.texto.delete("1.0", "end")
        self.texto.configure(state="disabled")
        self.hechos = 0
        self.ok_total = 0
        self.err_total = 0
        self.progreso.configure(maximum=len(trabajos), value=0)
        self._bloquea(True)
        self.b_abrir.configure(state="disabled")
        hilo = threading.Thread(target=self._trabajo, args=(trabajos, opts), daemon=True)
        hilo.start()

    def _trabajo(self, trabajos, opts):
        for i, (ruta, salida) in enumerate(trabajos, 1):
            self.cola.put(("inicio", i, len(trabajos), ruta))
            try:
                rep = procesa_archivo(ruta, salida, opts)
                self.cola.put(("listo", ruta, salida, rep))
            except zipfile.BadZipFile:
                self.cola.put(("error", ruta, "No es un archivo .xlsx válido o está dañado."))
            except PermissionError:
                self.cola.put(("error", ruta, "No se pudo escribir el resultado. Si el archivo de "
                                              "resultado está abierto en Excel, ciérrelo e intente de nuevo."))
            except Exception as e:  # noqa: BLE001
                self.cola.put(("error", ruta, "%s: %s" % (type(e).__name__, e)))
        self.cola.put(("fin",))

    def _revisar_cola(self):
        try:
            while True:
                msg = self.cola.get_nowait()
                tipo = msg[0]
                if tipo == "inicio":
                    _, i, total, ruta = msg
                    self.v_estado_txt.set("Procesando %d de %d: %s" % (i, total, os.path.basename(ruta)))
                elif tipo == "listo":
                    _, ruta, salida, rep = msg
                    self._muestra_reporte(ruta, salida, rep)
                    self.ultima_carpeta = os.path.dirname(salida)
                    self.ok_total += 1
                    self._avanza()
                elif tipo == "error":
                    _, ruta, texto = msg
                    self._escribe("ARCHIVO: %s" % os.path.basename(ruta), "titulo")
                    self._escribe("  ✗ No se pudo procesar. %s\n" % texto, "error")
                    self.err_total += 1
                    self._avanza()
                elif tipo == "fin":
                    self._bloquea(False)
                    resumen = "Proceso terminado: %d archivo(s) procesado(s)" % self.ok_total
                    if self.err_total:
                        resumen += ", %d con errores" % self.err_total
                    self.v_estado_txt.set(resumen + ".")
                    self._escribe(resumen + ".", "ok" if not self.err_total else "aviso")
                    if self.ultima_carpeta:
                        self.b_abrir.configure(state="normal")
        except queue.Empty:
            pass
        self.after(100, self._revisar_cola)

    def _avanza(self):
        self.hechos += 1
        self.progreso.configure(value=self.hechos)

    def _muestra_reporte(self, ruta, salida, rep):
        w = self._escribe
        w("ARCHIVO: %s" % os.path.basename(ruta), "titulo")
        w("  Guardado como: %s" % salida, "suave")
        w("  Hoja procesada: %s" % rep["hoja"], "suave")
        for a in rep["avisos"]:
            w("  ! %s" % a, "aviso")
        w("  ✓ Estados escritos:            %d" % rep["estado_escrito"], "ok")
        w("  ✓ Códigos de agencia ajustados: %d" % rep["agencia_reescrita"], "ok")
        w("  ✓ Teléfonos reescritos:         %d" % rep["telefono_reescrito"], "ok")
        if rep["estado_corregido"]:
            w("  ! Estados que tenían otro valor y se reemplazaron: %d" % len(rep["estado_corregido"]), "aviso")
            for n, a, b in rep["estado_corregido"]:
                w("      fila %d: %s → %s" % (n, a, b))
        if rep["codigo_invalido"]:
            w("  ✗ Códigos de agencia sin estado (columna L sin cambios): %d" % len(rep["codigo_invalido"]),
              "error")
            for n, guia, v in rep["codigo_invalido"]:
                w("      fila %d (%s): agencia %r" % (n, guia, v))
        if rep["telefono_invalido"]:
            w("  ! Teléfonos marcados con E (revisar): %d" % len(rep["telefono_invalido"]), "aviso")
            for n, guia, a, b in rep["telefono_invalido"]:
                w("      fila %d (%s): %s → %s" % (n, guia, a.replace("\n", " "), b))
        if rep["telefono_multiple"]:
            w("  ! Celdas con varios teléfonos (sin modificar): %d" % len(rep["telefono_multiple"]), "aviso")
            for n, guia, a in rep["telefono_multiple"]:
                w("      fila %d (%s): %s" % (n, guia, a.replace("\n", " | ")))
        w("")


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
