import math
import re
import unicodedata

import pandas as pd
from fpdf import FPDF

EXCEL_FILE = "excel1.xlsx"
PDF_FILE = "catalogo_hoteles.pdf"

# Controla si se incluye la portada (portada.jpg). Poner False para saltarla.
SHOW_PORTADA = False

# Controla si se incluye la página de presentación (Segunda-pagina.jpg).
# Poner True para volver a mostrarla; False para que el PDF empiece por el índice.
SHOW_SEGUNDA_PAGINA = False


def normalizar_provincia(nombre):
    """Normaliza provincia para ordenamiento alfabético sin tildes."""
    if not isinstance(nombre, str):
        nombre = str(nombre)
    nfkd = unicodedata.normalize("NFKD", nombre)
    sin_tildes = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sin_tildes.upper()


# Diccionario de capitales por provincia (claves normalizadas sin tildes)
CAPITALES = {
    "ACORUNA": "A Coruña",
    "ARABA": "Vitoria",
    "ALBACETE": "Albacete",
    "ALICANTE": "Alicante",
    "ALMERIA": "Almería",
    "ASTURIAS": "Oviedo",
    "AVILA": "Ávila",
    "BADAJOZ": "Badajoz",
    "BARCELONA": "Barcelona",
    "BURGOS": "Burgos",
    "CACERES": "Cáceres",
    "CADIZ": "Cádiz",
    "CANTABRIA": "Santander",
    "CASTELLON": "Castellón de la Plana",
    "CIUDADREAL": "Ciudad Real",
    "CORDOBA": "Córdoba",
    "CUENCA": "Cuenca",
    "CEUTA": "Ceuta",
    "GERONA": "Girona",
    "GRANADA": "Granada",
    "GUADALAJARA": "Guadalajara",
    "GUIPUZCOA": "San Sebastián",
    "HUELVA": "Huelva",
    "HUESCA": "Huesca",
    "ISLASBALEARES": "Palma de Mallorca",
    "JAEN": "Jaén",
    "LARIOJA": "Logroño",
    "LASPALMAS": "Las Palmas de Gran Canaria",
    "LEON": "León",
    "LLEIDA": "Lleida",
    "LUGO": "Lugo",
    "MADRID": "Madrid",
    "MALAGA": "Málaga",
    "MELILLA": "Melilla",
    "MURCIA": "Murcia",
    "NAVARRA": "Pamplona",
    "OURENSE": "Ourense",
    "PALENCIA": "Palencia",
    "PONTEVEDRA": "Pontevedra",
    "SALAMANCA": "Salamanca",
    "SEGOVIA": "Segovia",
    "SEVILLA": "Sevilla",
    "SORIA": "Soria",
    "TARRAGONA": "Tarragona",
    "TERUEL": "Teruel",
    "TOLEDO": "Toledo",
    "VALENCIA": "Valencia",
    "VALLADOLID": "Valladolid",
    "BIZKAIA": "Bilbao",
    "VIZCAYA": "Bilbao",
    "ZAMORA": "Zamora",
    "ZARAGOZA": "Zaragoza",
    "GIPUZKOA": "San Sebastián",
    "GIRONA": "Girona",
    "SANTACRUZDETENERIFE": "Santa Cruz de Tenerife",
    "TENERIFE": "Santa Cruz de Tenerife",
}

df = pd.read_excel(EXCEL_FILE)
df["CP"] = df["CP"].apply(lambda x: str(int(x)).zfill(5) if not pd.isnull(x) else "")
df = df.replace("?", "")

# Renombrar provincias para usar las denominaciones oficiales actuales
df["PROVINCIA"] = df["PROVINCIA"].replace({"ÁLAVA": "ARABA"})


# Extraer valor numérico de la clasificación para ordenar por estrellas (5->0)
def extraer_estrellas(val):
    try:
        s = str(val)
        m = re.search(r"(\d+)", s)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return 0


df["ESTRELLAS"] = df["CLASIFICACION HOTEL"].apply(extraer_estrellas)


# Crear función de normalización robusta para localidades/capitales
def normalizar_ciudad(nombre):
    if not isinstance(nombre, str):
        nombre = str(nombre)
    nfkd = unicodedata.normalize("NFKD", nombre)
    sin_tildes = "".join(c for c in nfkd if not unicodedata.combining(c))
    s = sin_tildes.replace("/", " ").replace("-", " ")
    s = " ".join(s.split())
    return s.upper()


# Crear columna ES_CAPITAL: True si la localidad es la capital de su provincia
def es_capital(row):
    provincia_norm = normalizar_provincia(row.get("PROVINCIA", "")).replace(" ", "")
    capital_oficial = CAPITALES.get(provincia_norm, "")
    localidad = str(row.get("LOCALIDAD", "")).strip()
    if not capital_oficial or not localidad:
        return False

    cap_norm = normalizar_ciudad(capital_oficial)

    # Si la localidad contiene variantes separadas por '/' o '-', comparar
    # cada variante de forma independiente usando igualdad exacta tras
    # normalizar. Esto evita falsos positivos por contención parcial.
    if "/" in localidad or "-" in localidad:
        partes = re.split(r"[/-]", localidad)
        for p in partes:
            if not p:
                continue
            if normalizar_ciudad(p) == cap_norm:
                return True
        return False

    # Localidad sin variantes: comparar igualdad exacta tras normalizar
    loc_norm = normalizar_ciudad(localidad)
    return loc_norm == cap_norm


df["ES_CAPITAL"] = df.apply(es_capital, axis=1)

# Columna auxiliar para ordenar por nombre limpio (sin "HOTEL" al inicio, sin tildes)
def _nombre_orden(x):
    s = str(x).strip()
    if s.upper().startswith("HOTEL"):
        s = s[5:].strip()
    if s.upper().endswith("S.L."):
        s = s[:-4].strip()
    elif s.upper().endswith("S.L"):
        s = s[:-3].strip()
    return normalizar_provincia(s)

df["NOMBRE_ORDEN"] = df["NOMBRE DE EMPRESA"].apply(_nombre_orden)

# Ordenar por: provincia → ES_CAPITAL (True primero) → localidad → estrellas descendentes → nombre alfabético
df = df.sort_values(
    by=["PROVINCIA", "ES_CAPITAL", "LOCALIDAD", "ESTRELLAS", "NOMBRE_ORDEN"],
    key=lambda col: (
        col.map(normalizar_provincia) if col.name in ["PROVINCIA", "LOCALIDAD"] else col
    ),
    ascending=[True, False, True, False, True],
)


# --- PDF ---
class PDF(FPDF):
    def __init__(self):
        # Tamaño de página nativo 6" x 9" (KDP paperback)
        super().__init__(orientation="P", unit="mm", format=(PAGE_WIDTH, PAGE_HEIGHT))
        self.set_margins(MARGIN_GUTTER, Y_TOP, MARGIN_OUTER + BLEED)

    def header(self):
        # Márgenes simétricos: el medianil cambia de lado en cada página
        izq, der = margenes_pagina(self.page_no())
        self.set_margins(izq, Y_TOP, der)
        self.set_xy(izq, Y_TOP)

        if getattr(self, "provincia_actual", "") in [None, "", False]:
            return

        # Encabezado por provincia
        self.set_font("Helvetica", "B", FONT_CABECERA)
        self.set_text_color(*AZUL_PORTADA)
        # Agregar "(cont)" si esta es una página de continuación
        provincia_text = f"PROVINCIA DE {self.provincia_actual.upper()}"
        if getattr(self, "provincia_continuacion", False):
            provincia_text += " (cont.)"
        self.cell(
            0,
            5,
            _enc(provincia_text),
            new_x="LMARGIN", new_y="NEXT",
            align="C",
        )

        # Línea superior decorativa
        self.set_draw_color(180, 180, 180)
        self.set_line_width(0.2)
        self.line(izq, Y_LINEA, PAGE_WIDTH - der, Y_LINEA)
        self.set_text_color(0, 0, 0)

    def footer(self):
        if getattr(self, "provincia_actual", "") is None:
            return
        # Las portadas azules llevan su propio número (blanco, dentro del panel)
        if self.page_no() in getattr(self, "paginas_sin_pie", set()):
            return
        izq = x_contenido(self.page_no())
        self.set_xy(izq, Y_PIE)
        self.set_font("Helvetica", "I", 7)
        self.set_text_color(128)
        self.cell(0, 4.5, f"{self.page_no()}", align="C")
        self.set_text_color(0, 0, 0)


# --- Función para calcular altura real de UNA LÍNEA ---
# Factor de seguridad: el ajuste de línea real corta por palabras, así que una
# línea puede ocupar más alto que el que da la división ancho_texto/ancho_total.
# Con columnas estrechas (6"x9") esto es crítico para no salirse del margen.
FACTOR_SEGURIDAD_ANCHO = 0.90


def calcular_altura_linea(pdf, texto, ancho_efectivo, alto_linea):
    """Calcula cuántas líneas ocupa un texto dado el ancho disponible."""
    if not texto:
        return 0
    w = pdf.get_string_width(texto)
    num_lineas = max(1, math.ceil(w / (ancho_efectivo * FACTOR_SEGURIDAD_ANCHO)))
    return num_lineas * alto_linea


# --- Función para calcular altura real del bloque completo ---
def calcular_altura_bloque(pdf, lineas_list, ancho_efectivo, alto_linea):
    """Calcula la altura total de un bloque con varias líneas."""
    total_altura = 2  # pequeño margen al inicio
    for linea in lineas_list:
        if linea:
            w = pdf.get_string_width(linea)
            num_lineas = max(1, math.ceil(w / (ancho_efectivo * FACTOR_SEGURIDAD_ANCHO)))
            total_altura += num_lineas * alto_linea
    total_altura += 2  # pequeño margen al final
    return total_altura


def limpiar_nombre_hotel(nombre):
    """Elimina la palabra 'HOTEL' al inicio y 'S.L.' al final si existen."""
    nombre_clean = str(nombre).strip()
    if nombre_clean.upper().startswith("HOTEL"):
        nombre_clean = nombre_clean[5:].strip()
    if nombre_clean.upper().endswith("S.L."):
        nombre_clean = nombre_clean[:-4].strip()
    elif nombre_clean.upper().endswith("S.L"):
        nombre_clean = nombre_clean[:-3].strip()
    return nombre_clean


# Palabras que deben ir en minúscula cuando aparecen en medio de una dirección
_PREPOSICIONES = {
    "De", "Del", "O", "Y", "A", "E", "En", "Con", "Por", "Para", "Sin",
    "La", "Las", "Los", "El", "Al",
}


def corregir_preposiciones(texto):
    """Aplica .title() y luego pone en minúscula las preposiciones/conjunciones
    que aparezcan en posición no inicial dentro del texto."""
    texto = str(texto).title()
    palabras = texto.split()
    resultado = []
    for i, palabra in enumerate(palabras):
        # La primera palabra siempre en Title Case
        if i == 0:
            resultado.append(palabra)
        elif palabra in _PREPOSICIONES:
            resultado.append(palabra.lower())
        else:
            resultado.append(palabra)
    return " ".join(resultado)


# ---------------------------------------------------------------------------
# TAMAÑO DE PÁGINA Y MÁRGENES (KDP paperback 6" x 9" CON SANGRADO)
# ---------------------------------------------------------------------------
# Tamaño final del libro (trim): 6" x 9" = 152.4 x 228.6 mm.
#
# Como las portadas azules de sección llevan el fondo hasta el borde del papel,
# el interior se genera CON SANGRADO (bleed). KDP exige entonces que el PDF
# mida trim + 0.125" (3.175 mm) por arriba, por abajo y por el borde EXTERIOR
# (el lado del lomo no lleva sangrado):
#       155.575 x 234.95 mm  =  6.125" x 9.25"
# IMPORTANTE: al subirlo hay que marcar la opción "Con sangrado / Bleed".
#
# Márgenes (medidos desde el corte final) exigidos por KDP para 501-828 págs:
#   - Medianil (margen interior, junto al lomo): 0.875" = 22.23 mm
#   - Margen exterior / superior / inferior:     mínimo 0.25" = 6.35 mm
# Usamos 0.5" (12.7 mm) en exterior/superior/inferior por seguridad.
#
# El medianil alterna de lado: en páginas impares (derechas) el interior es el
# borde izquierdo; en páginas pares (izquierdas), el derecho. El sangrado va
# siempre en el borde contrario al lomo.
# ---------------------------------------------------------------------------
TRIM_WIDTH = 152.4
TRIM_HEIGHT = 228.6
BLEED = 3.175           # 0.125"

# Tamaño físico del PDF (incluye el sangrado)
PAGE_WIDTH = TRIM_WIDTH + BLEED
PAGE_HEIGHT = TRIM_HEIGHT + 2 * BLEED

# Con ~594 páginas KDP exige un medianil mínimo de 0.75" (19.05 mm). Usamos
# 19.3 mm de interior y 17.1 mm de exterior: es el reparto más simétrico
# posible sin incumplir ese mínimo, de modo que la mancha se ve prácticamente
# centrada en la página (solo 2.2 mm de diferencia entre ambos lados).
MARGIN_GUTTER = 19.3    # medianil (interior, junto al lomo)
MARGIN_OUTER = 17.1     # margen exterior
MARGIN_TOP = 12.7
MARGIN_BOTTOM = 12.7

CONTENT_WIDTH = TRIM_WIDTH - MARGIN_GUTTER - MARGIN_OUTER
CONTENT_HEIGHT = TRIM_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM

# Coordenadas verticales de la mancha, ya sobre el papel con sangrado
Y_TOP = BLEED + MARGIN_TOP
Y_BOTTOM = PAGE_HEIGHT - BLEED - MARGIN_BOTTOM

# Compatibilidad con el resto del código (márgenes "simétricos" de referencia)
MARGIN = MARGIN_OUTER


def margenes_pagina(page_no):
    """Devuelve (margen_izquierdo, margen_derecho) sobre el papel CON sangrado.

    Impar = página derecha (recto) → el interior (lomo) es el borde izquierdo,
            así que el sangrado se suma al margen derecho.
    Par   = página izquierda (verso) → al revés.
    """
    if page_no % 2 == 1:
        return MARGIN_GUTTER, MARGIN_OUTER + BLEED
    return MARGIN_OUTER + BLEED, MARGIN_GUTTER


def x_contenido(page_no):
    """Coordenada X donde empieza el área de texto en esa página."""
    return margenes_pagina(page_no)[0]


def x_corte(page_no):
    """Coordenada X donde empieza el área que quedará tras el corte (trim)."""
    return 0.0 if page_no % 2 == 1 else BLEED


# CONFIGURACIÓN DE GRID FLEXIBLE
# Las columnas ocupan EXACTAMENTE la mancha: el hueco de separación va solo
# ENTRE columnas, nunca antes de la primera ni después de la última, para que
# el bloque de texto quede perfectamente ajustado a los márgenes.
COLS = 3
SEP_COLUMNAS = 3.5
COLUMN_WIDTH = (CONTENT_WIDTH - (COLS - 1) * SEP_COLUMNAS) / COLS
PASO_COLUMNA = COLUMN_WIDTH + SEP_COLUMNAS

# Cabecera de provincia + línea decorativa
Y_LINEA = Y_TOP + 6.5
Y_START = Y_TOP + 9.5
# El número de página del pie se imprime justo encima del margen inferior
Y_PIE = Y_BOTTOM - 4.5
Y_LIMIT = Y_PIE - 1.0

# Tipografías del catálogo (ajustadas al ancho real de columna de 6"x9"
# y a la densidad necesaria para mantener el libro por debajo de 600 páginas)
FONT_CABECERA = 9.5
FONT_LOCALIDAD = 6.7
FONT_NOMBRE = 6.1
FONT_CAT = 5.5
FONT_DETALLE = 5.5

line_height = 2.8
# Pequeño colchón para que ninguna línea toque el borde de la mancha
ancho_texto = COLUMN_WIDTH - 1.5

# --- PALETA DE COLOR (tono de las fotos, ligeramente hacia el cian) ---
AZUL_PORTADA = (64, 152, 193)   # fondo de las portadas azules
AZUL_ACENTO = (64, 152, 193)    # todos los azules usan el mismo tono


def formatear_clasificacion(val):
    """Convierte '5 *' → '5*'. El resto (LLAVES, ESPIGAS, CATEGORÍA) se deja igual."""
    s = str(val).strip()
    m = re.match(r"^(\d+)\s*\*$", s)
    if m:
        return f"{m.group(1)}*"
    return s


def _enc(s):
    """Codifica a latin-1 (fuentes core de FPDF)."""
    return str(s).encode("latin-1", "ignore").decode("latin-1")


def construir_lineas_hotel(row):
    """Construye el diccionario de líneas de un hotel con el formato de las fotos.

    Orden de impresión: categoría → NOMBRE (negrita) → registro → dirección →
    CP+localidad → teléfono → web.
    """
    _clasif = str(row["CLASIFICACION HOTEL"]).strip()
    _hab = str(row["NRO. HABITACIONES"]).strip()
    _clasif_ok = _clasif not in ("", "-", "nan", "NaN", "?")
    _hab_ok = _hab not in ("", "-", "nan", "NaN", "?") and _hab.replace(".", "").isdigit()

    partes = []
    if _clasif_ok:
        partes.append(formatear_clasificacion(_clasif))
    if _hab_ok:
        partes.append(f"{_hab} hab.")
    # Modalidad: siempre precedida de "hotel" (p.ej. "hotel playa"), como en las fotos
    _mod = str(row.get("MODALIDAD", "")).strip()
    _mod_ok = _mod not in ("", "-", "nan", "NaN", "?", "None")
    if _mod_ok:
        partes.append(f"Hotel {_mod.lower()}")
    linea_cat = " - ".join(partes)

    linea_nombre = limpiar_nombre_hotel(row["NOMBRE DE EMPRESA"]).upper()

    _reg = str(row.get("N. REGISTRO", "")).strip()
    _reg_ok = _reg not in ("", "-", "nan", "NaN", "?", "None")
    linea_reg = f"Registro oficial: {_reg}" if _reg_ok else ""

    _dir = str(row["DIRECCION"]).strip()
    linea_dir = corregir_preposiciones(_dir) if _dir not in ("", "-", "nan", "NaN", "?") else ""

    linea_loc = corregir_preposiciones(f"{row['CP']} {row['LOCALIDAD']}")

    _tel = str(row["TELEFONO1"]).strip()
    linea_tel = f"Tel. {_tel}" if _tel not in ("", "-", "nan", "NaN", "?") else ""

    _web = str(row["SITIO WEB"]).strip()
    linea_web = f"Web: {_web.lower()}" if _web not in ("", "-", "nan", "NaN", "?") else ""

    return {
        "cat": _enc(linea_cat),
        "nombre": _enc(linea_nombre),
        "reg": _enc(linea_reg),
        "dir": _enc(linea_dir),
        "loc": _enc(linea_loc),
        "tel": _enc(linea_tel),
        "web": _enc(linea_web),
    }


# ---------------------------------------------------------------------------
# PORTADAS AZULES DE SECCIÓN (estilo de las fotos)
# ---------------------------------------------------------------------------
def _dibujar_separador(pdf, x0, ancho, y):
    """Separador decorativo blanco: línea — rombo — línea, centrado en `y`."""
    pdf.set_draw_color(255, 255, 255)
    pdf.set_fill_color(255, 255, 255)
    pdf.set_line_width(0.5)
    cx = x0 + ancho / 2
    largo = ancho * 0.28   # longitud de cada media línea
    hueco = ancho * 0.05   # separación entre línea y rombo
    pdf.line(cx - hueco - largo, y, cx - hueco, y)
    pdf.line(cx + hueco, y, cx + hueco + largo, y)
    # Rombo central (4 vértices)
    s = 2.0
    pdf.polygon(
        [(cx, y - s), (cx + s, y), (cx, y + s), (cx - s, y)],
        style="F",
    )


def _tamano_fuente_ajustado(pdf, lineas, ancho_util, size_max=13, size_min=6):
    """Mayor tamaño de fuente (Helvetica Bold) con el que TODAS las líneas
    caben en `ancho_util`."""
    size = size_max
    while size > size_min:
        pdf.set_font("Helvetica", "B", size)
        if all(pdf.get_string_width(_enc(l)) <= ancho_util for l in lineas):
            return size
        size -= 0.5
    return size_min


def dibujar_portada_seccion(pdf, lineas_es, lineas_en, page_number_display):
    """Portada azul de sección, bilingüe, al estilo de las fotos:
    - Azul a sangre: cubre TODO el papel, sangrado incluido, para que no quede
      ningún borde blanco después del corte
    - Bloque en español (mitad superior) centrado
    - Separador decorativo (línea — rombo — línea)
    - Bloque en inglés (mitad inferior) centrado
    - Número de página blanco arriba a la derecha
    """
    # El pie se dibuja al cerrar la página, cuando `provincia_actual` puede
    # haber cambiado ya; marcamos la página para que no lo imprima encima.
    if not hasattr(pdf, "paginas_sin_pie"):
        pdf.paginas_sin_pie = set()
    pdf.paginas_sin_pie.add(pdf.page_no())

    # Fondo azul a sangre completa (incluida la zona de sangrado)
    pdf.set_fill_color(*AZUL_PORTADA)
    pdf.rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT, "F")
    pdf.set_text_color(255, 255, 255)

    # El texto se centra respecto al ÁREA DE CORTE (lo que queda del papel tras
    # guillotinar), no respecto a la mancha: así se ve perfectamente centrado.
    # El ancho útil se limita al medianil por ambos lados, de modo que el
    # bloque centrado sigue respetando el margen interior.
    x0 = x_corte(pdf.page_no())
    y0 = BLEED
    ancho = TRIM_WIDTH
    alto = TRIM_HEIGHT

    # Número de página arriba a la derecha, dentro de los márgenes
    pdf.set_font("Helvetica", "", 9)
    pdf.set_xy(x_contenido(pdf.page_no()) + CONTENT_WIDTH - 20, Y_TOP + 4)
    pdf.cell(15, 6, str(page_number_display), align="R")

    ancho_util = ancho - 2 * MARGIN_GUTTER
    size = _tamano_fuente_ajustado(pdf, lineas_es + lineas_en, ancho_util)
    alt = size * 0.62  # alto de línea proporcional al cuerpo

    # Bloque español (centrado alrededor del 30% de la mancha)
    pdf.set_font("Helvetica", "B", size)
    y_es = y0 + alto * 0.30 - (len(lineas_es) * alt) / 2
    pdf.set_xy(x0, y_es)
    for linea in lineas_es:
        pdf.set_x(x0)
        pdf.cell(ancho, alt, _enc(linea), align="C", new_x="LEFT", new_y="NEXT")

    # Separador decorativo en el centro vertical de la mancha
    _dibujar_separador(pdf, x0, ancho, y0 + alto * 0.505)

    # Bloque inglés (centrado alrededor del 68% de la mancha)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", size)
    y_en = y0 + alto * 0.68 - (len(lineas_en) * alt) / 2
    pdf.set_xy(x0, y_en)
    for linea in lineas_en:
        pdf.set_x(x0)
        pdf.cell(ancho, alt, _enc(linea), align="C", new_x="LEFT", new_y="NEXT")

    # Resetear estilo
    pdf.set_text_color(0, 0, 0)
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.2)


# --- TEXTOS DE LAS PORTADAS AZULES (exactos de las fotos) ---
PORTADA_CATALOGO_ES = [
    "GUIA DE HOTELES DE ESPAÑA",
    "ORDENADOS POR PROVINCIAS.",
    "EN CADA PROVINCIA POR ORDEN",
    "DE CATEGORÍA (ESTRELLAS).",
    "Y DENTRO DE LA MISMA CATEGORÍA",
    "DETALLADOS POR ORDEN ALFABÉTICO",
    "DEL NOMBRE DEL HOTEL",
]
PORTADA_CATALOGO_EN = [
    "A GUIDE TO HOTELS IN SPAIN",
    "ORGANIZED BY PROVINCES.",
    "WITHIN EACH PROVINCE, BY CATEGORY",
    "(STARS). AND WITHIN EACH CATEGORY,",
    "LISTED IN ALPHABETICAL ORDER",
    "BY HOTEL NAME",
]
PORTADA_HOTELES_ES = [
    "NOMBRE DE LOS HOTELES",
    "DE ESTA GUÍA, POR ORDEN",
    "ALFABÉTICO Y NÚMERO DE",
    "LA PÁGINA DONDE SE",
    "ENCUENTRAN.",
]
PORTADA_HOTELES_EN = [
    "NAMES OF THE HOTELS",
    "IN THIS GUIDE, IN ALPHABETICAL",
    "ORDER AND THE PAGE NUMBER",
    "WHERE THEY ARE LOCATED.",
]
PORTADA_POBLACIONES_ES = [
    "CIUDADES Y LOCALIDADES",
    "DE ESPAÑA CON HOTELES, POR",
    "ORDEN ALFABÉTICO Y NÚMERO",
    "DE LA PÁGINA DONDE SE",
    "ENCUENTRAN.",
]
PORTADA_POBLACIONES_EN = [
    "CITIES AND TOWNS OF SPAIN",
    "WITH HOTELS, IN ALPHABETICAL",
    "ORDER AND THE PAGE NUMBER",
    "WHERE THEY ARE LOCATED.",
]


# Obtener lista única de provincias en orden alfabético (sin tildes)
provincias_unicas = sorted(df["PROVINCIA"].unique().tolist(), key=normalizar_provincia)

# Estructura para guardar índice de provincias y sus páginas
indice_provincias = []

# Rellenar indice_provincias con capitales del diccionario
for prov in provincias_unicas:
    prov_normalizada = normalizar_provincia(prov).replace(" ", "")
    capital = CAPITALES.get(prov_normalizada, "-")
    indice_provincias.append({"provincia": prov, "capital": capital, "pagina": None})

# ---------------------------------------------------------------------------
# ESTRATEGIA DE DOBLE RENDER (índices 100% exactos)
# ---------------------------------------------------------------------------
# El índice de provincias necesita los números de página REALES del catálogo,
# pero debe aparecer ANTES del catálogo en el PDF. En vez de *estimar* las
# alturas (lo que desincronizaba el índice del PDF real), renderizamos el
# catálogo DOS VECES con exactamente el mismo código:
#   Pasada 1 → a un PDF temporal, solo para capturar las páginas reales.
#   Pasada 2 → al PDF final, ya con los números de página correctos.
# Como el render es idéntico y va precedido del mismo nº de páginas fijas,
# la paginación coincide al 100%.
# ---------------------------------------------------------------------------


def render_catalogo(pdf):
    """Dibuja TODO el catálogo por provincias en `pdf`.

    Devuelve (prov_pages, hotel_pages, loc_pages): la página REAL de la primera
    aparición de cada provincia, hotel (nombre limpio) y localidad.
    """
    prov_pages = {}
    hotel_pages = {}
    loc_pages = {}

    def columnas(page_no):
        """Coordenadas X de las 3 columnas en esa página (el medianil alterna)."""
        base = x_contenido(page_no)
        return [base + i * PASO_COLUMNA for i in range(COLS)]

    x_positions = columnas(1)
    pdf.provincia_actual = ""
    y_actual = [Y_START] * COLS
    provincia_anterior = ""
    localidad_anterior = ""
    current_col = 0

    for idx, row in df.iterrows():
        provincia = str(row["PROVINCIA"])
        localidad = str(row["LOCALIDAD"])
        hotel_name = str(row["NOMBRE DE EMPRESA"]).strip()

        # CAMBIO DE PROVINCIA → NUEVA PÁGINA Y RESET DE ALTURAS
        if provincia != provincia_anterior:
            provincia_anterior = provincia
            localidad_anterior = ""
            pdf.provincia_actual = provincia
            pdf.provincia_continuacion = False
            pdf.add_page()
            x_positions = columnas(pdf.page_no())
            current_col = 0
            y_actual = [Y_START] * COLS
            if provincia not in prov_pages:
                prov_pages[provincia] = pdf.page_no()

        hotel_name_display = limpiar_nombre_hotel(hotel_name)

        _d = construir_lineas_hotel(row)
        linea_cat = _d["cat"]
        linea_nombre = _d["nombre"]
        linea_reg = _d["reg"]
        linea_dir = _d["dir"]
        linea_loc = _d["loc"]
        linea_tel = _d["tel"]
        linea_web = _d["web"]
        lineas_hotel = [
            linea_nombre, linea_cat, linea_reg,
            linea_dir, linea_loc, linea_tel, linea_web,
        ]

        # Altura estimada del hotel (solo para decidir salto de columna/página)
        pdf.set_font("Helvetica", "", FONT_NOMBRE)
        altura_hotel = calcular_altura_bloque(
            pdf, [_l for _l in lineas_hotel if _l], ancho_texto, line_height
        )

        hay_cambio_localidad = localidad != localidad_anterior
        altura_localidad = 0
        if hay_cambio_localidad:
            altura_localidad = (
                calcular_altura_linea(pdf, localidad.upper(), COLUMN_WIDTH, line_height) + 4
            )

        altura_total_requerida = altura_localidad + altura_hotel + 2
        localidad_cont = False

        if y_actual[current_col] + altura_total_requerida > Y_LIMIT:
            current_col += 1
            if current_col >= COLS:
                pdf.provincia_continuacion = True
                if not hay_cambio_localidad:
                    localidad_cont = True
                pdf.add_page()
                x_positions = columnas(pdf.page_no())
                current_col = 0
                y_actual = [Y_START] * COLS

        x = x_positions[current_col]
        y_pos = y_actual[current_col]

        # ---- REGISTRAR HOTEL CON SU PÁGINA REAL (ya resuelto el salto de página) ----
        if hotel_name_display and hotel_name_display not in hotel_pages:
            hotel_pages[hotel_name_display] = pdf.page_no()

        # TÍTULO DE LOCALIDAD
        if hay_cambio_localidad:
            y_pos = y_pos + 1
            localidad_anterior = localidad
            if localidad not in loc_pages:
                loc_pages[localidad] = pdf.page_no()
            pdf.set_xy(x, y_pos)
            pdf.set_font("Helvetica", "B", FONT_LOCALIDAD)
            pdf.set_text_color(*AZUL_ACENTO)
            pdf.multi_cell(COLUMN_WIDTH, line_height, _enc(localidad.upper()), border=0, align="L")
            y_pos = pdf.get_y()
            y_actual[current_col] = y_pos
        elif localidad_cont:
            y_pos = y_pos + 1
            pdf.set_xy(x_positions[0], y_pos)
            pdf.set_font("Helvetica", "B", FONT_LOCALIDAD)
            pdf.set_text_color(*AZUL_ACENTO)
            pdf.multi_cell(COLUMN_WIDTH, line_height, _enc(localidad.upper() + " (cont.)"), border=0, align="L")
            cont_y = pdf.get_y()
            for _c in range(COLS):
                y_actual[_c] = cont_y
            y_pos = cont_y
            x = x_positions[current_col]

        # TEXTO DEL HOTEL
        pdf.set_xy(x, y_pos)
        pdf.set_text_color(0, 0, 0)
        if linea_cat:
            pdf.set_font("Helvetica", "B", FONT_CAT)
            pdf.multi_cell(ancho_texto, line_height, linea_cat, border=0, align="L")
        pdf.set_x(x)
        pdf.set_font("Helvetica", "B", FONT_NOMBRE)
        pdf.multi_cell(ancho_texto, line_height, linea_nombre, border=0, align="L")
        pdf.set_font("Helvetica", "", FONT_DETALLE)
        if linea_reg:
            pdf.set_x(x)
            pdf.multi_cell(ancho_texto, line_height, linea_reg, border=0, align="L")
        if linea_dir:
            pdf.set_x(x)
            pdf.multi_cell(ancho_texto, line_height, linea_dir, border=0, align="L")
        pdf.set_x(x)
        pdf.multi_cell(ancho_texto, line_height, linea_loc, border=0, align="L")
        if linea_tel:
            pdf.set_x(x)
            pdf.multi_cell(ancho_texto, line_height, linea_tel, border=0, align="L")
        if linea_web:
            pdf.set_x(x)
            pdf.multi_cell(ancho_texto, line_height, linea_web, border=0, align="L")

        y_actual[current_col] = pdf.get_y() + 2

    return prov_pages, hotel_pages, loc_pages


# ---- PASADA 1: render de medición (a un PDF temporal) ----
# Páginas fijas antes del catálogo: [portada opc.] + [intro opc.] + índice + portada azul.
paginas_fijas_antes = (
    (1 if SHOW_PORTADA else 0)
    + (1 if SHOW_SEGUNDA_PAGINA else 0)
    + 2  # índice de provincias + portada azul del catálogo
)

_scratch = PDF()
_scratch.set_auto_page_break(auto=False)
_scratch.set_font("Helvetica", "", 9)
_scratch.provincia_actual = None  # sin cabecera/pie en las páginas fijas dummy
for _ in range(paginas_fijas_antes):
    _scratch.add_page()
prov_pages_real, _hotel_pages_m, _loc_pages_m = render_catalogo(_scratch)
del _scratch

# Índice de provincias con las páginas REALES
for item in indice_provincias:
    prov = item["provincia"]
    if prov in prov_pages_real:
        item["pagina"] = prov_pages_real[prov]

# ---- PASADA 2: generar el PDF completo en orden correcto ----

# --- CREAR PDF FINAL ---
pdf = PDF()
pdf.set_auto_page_break(auto=False)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(0, 0, 0)
pdf.provincia_continuacion = False

# Añadir portada a toda la página si existe
if SHOW_PORTADA:
    try:
        pdf.add_page()
        PAGE_W = pdf.w
        PAGE_H = pdf.h
        pdf.image("portada.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
    except Exception as e:
        print(f"No se pudo cargar portada.jpg: {e}")

# Añadir página de presentación (Segunda-pagina.jpg) solo si está activada
if SHOW_SEGUNDA_PAGINA:
    try:
        pdf.add_page()
        PAGE_W = pdf.w
        PAGE_H = pdf.h
        pdf.image("Segunda-pagina.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
    except Exception as e:
        print(f"No se pudo cargar Segunda-pagina.jpg: {e}")

# --- PÁGINA DE ÍNDICE 1: PROVINCIAS Y SUS CAPITALES ---
pdf.provincia_actual = None
pdf.add_page()

X_IDX = x_contenido(pdf.page_no())

# Número de página arriba a la derecha (estilo foto)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(0, 0, 0)
pdf.set_xy(X_IDX + CONTENT_WIDTH - 15, Y_TOP)
pdf.cell(15, 6, str(pdf.page_no()), align="R")

# Cabecera "ÍNDICE 1  -  INDEX 1"
pdf.set_xy(X_IDX, Y_TOP + 1)
pdf.set_font("Helvetica", "B", 10)
pdf.cell(CONTENT_WIDTH, 6, _enc("ÍNDICE 1     -     INDEX 1"), align="C", new_x="LEFT", new_y="NEXT")
pdf.ln(1)
pdf.set_font("Helvetica", "B", 11)
pdf.cell(CONTENT_WIDTH, 6, _enc("PROVINCIAS DE ESPAÑA Y SUS CAPITALES"), new_x="LEFT", new_y="NEXT", align="C")
pdf.set_font("Helvetica", "B", 9)
pdf.cell(CONTENT_WIDTH, 5, "PROVINCES OF SPAIN AND THEIR CAPITALS", new_x="LEFT", new_y="NEXT", align="C")
pdf.ln(3)

usable_width_prov = CONTENT_WIDTH
separation_prov = 5
table_width_prov = (usable_width_prov - separation_prov) / 2
col_widths_prov = [table_width_prov * 0.41, table_width_prov * 0.45, table_width_prov * 0.14]
x_left_prov = X_IDX
x_right_prov = X_IDX + table_width_prov + separation_prov

n_prov = len(indice_provincias)
mid_prov = (n_prov + 1) // 2
left_items_prov = indice_provincias[:mid_prov]
right_items_prov = indice_provincias[mid_prov:]
while len(left_items_prov) < len(right_items_prov):
    left_items_prov.append({"provincia": "", "capital": "", "pagina": None})
while len(right_items_prov) < len(left_items_prov):
    right_items_prov.append({"provincia": "", "capital": "", "pagina": None})

# Alto de fila calculado para repartir las provincias por toda la página
_alto_disp_prov = Y_LIMIT - pdf.get_y()
row_h_prov = min(7.0, _alto_disp_prov / (len(left_items_prov) + 1))

pdf.set_font("Helvetica", "B", 7)
y_header_prov = pdf.get_y()
for _x_tabla in (x_left_prov, x_right_prov):
    pdf.set_xy(_x_tabla, y_header_prov)
    pdf.cell(col_widths_prov[0], row_h_prov, "PROVINCIAS", border=1, align="C")
    pdf.cell(col_widths_prov[1], row_h_prov, "CAPITALES", border=1, align="C")
    pdf.cell(col_widths_prov[2], row_h_prov, _enc("Pág."), border=1, align="C")
pdf.set_y(y_header_prov + row_h_prov)

# Helper: imprime una celda ajustando el tamaño de fuente si el texto
# no cabe en el ancho disponible. Empieza en `font_size_default` y baja
# hasta `font_size_min` en pasos de 0.5 hasta encontrar uno que quepa
# (con un pequeño padding interno). Si ni al mínimo cabe, usa el mínimo.
def cell_ajustada(pdf, w, h, txt, align, font_family="Helvetica", font_style="",
                  font_size_default=7, font_size_min=4.5, padding=1.0):
    txt_safe = txt.encode("latin-1", "ignore").decode("latin-1")
    ancho_util = w - padding * 2
    size = font_size_default
    while size >= font_size_min:
        pdf.set_font(font_family, font_style, size)
        if pdf.get_string_width(txt_safe) <= ancho_util:
            break
        size -= 0.5
    pdf.cell(w, h, txt_safe, border=1, align=align)
    # Restaurar tamaño por defecto para celdas siguientes
    pdf.set_font(font_family, font_style, font_size_default)


for i in range(len(left_items_prov)):
    left_p = left_items_prov[i]
    right_p = right_items_prov[i] if i < len(right_items_prov) else {"provincia": "", "capital": "", "pagina": None}
    prov_l = left_p["provincia"]
    prov_r = right_p["provincia"]
    if not prov_l and not prov_r:
        continue
    y_p = pdf.get_y()
    capital_l = left_p["capital"]
    page_l = str(left_p["pagina"]) if left_p["pagina"] is not None else "..."

    # FILA IZQUIERDA
    pdf.set_xy(x_left_prov, y_p)
    cell_ajustada(pdf, col_widths_prov[0], row_h_prov, prov_l, "L")
    cell_ajustada(pdf, col_widths_prov[1], row_h_prov, capital_l, "L")
    cell_ajustada(pdf, col_widths_prov[2], row_h_prov, page_l, "C")

    capital_r = right_p["capital"]
    page_r = str(right_p["pagina"]) if right_p["pagina"] is not None else "..."

    # FILA DERECHA
    pdf.set_xy(x_right_prov, y_p)
    cell_ajustada(pdf, col_widths_prov[0], row_h_prov, prov_r, "L")
    cell_ajustada(pdf, col_widths_prov[1], row_h_prov, capital_r, "L")
    cell_ajustada(pdf, col_widths_prov[2], row_h_prov, page_r, "C")

    pdf.set_y(y_p + row_h_prov)

# --- PORTADA AZUL DEL CATÁLOGO (antes de las provincias) ---
pdf.provincia_actual = None
pdf.add_page()
dibujar_portada_seccion(
    pdf,
    PORTADA_CATALOGO_ES,
    PORTADA_CATALOGO_EN,
    pdf.page_no(),
)

# --- GENERAR CATÁLOGO (pasada 2, render final; páginas idénticas a la pasada 1) ---
prov_pages_final, hotel_pages, loc_pages = render_catalogo(pdf)

# --- PORTADA ÍNDICE ALFABÉTICO DE HOTELES (estilo minimalista) ---
pdf.provincia_actual = None
pdf.add_page()
dibujar_portada_seccion(
    pdf,
    PORTADA_HOTELES_ES,
    PORTADA_HOTELES_EN,
    pdf.page_no(),
)

# --- INICIAR ÍNDICE ALFABÉTICO DE HOTELES ---
pdf.provincia_actual = None
pdf.add_page()


# --- Cabecera común de las páginas de índice alfabético ---
# Los índices finales van muy compactos (4 columnas) para no inflar el
# número total de páginas del libro.
FONT_TITULO_INDICE = 7.5
FONT_INDICE = 5.0
ROW_H_INDICE = 2.9
COLS_INDICE = 4
SEP_INDICE = 2.5
Y_LIMIT_INDICE = Y_LIMIT


def cabecera_indice(pdf, titulo_es, titulo_en):
    """Imprime los dos títulos bilingües y deja el cursor bajo ellos."""
    x = x_contenido(pdf.page_no())
    pdf.set_xy(x, Y_TOP)
    pdf.set_font("Helvetica", "B", FONT_TITULO_INDICE)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(CONTENT_WIDTH, 4.5, _enc(titulo_es), new_x="LEFT", new_y="NEXT", align="C")
    pdf.cell(CONTENT_WIDTH, 4.5, _enc(titulo_en), new_x="LEFT", new_y="NEXT", align="C")
    pdf.ln(1.5)
    return pdf.get_y()


TITULO_HOTELES_ES = "Hoteles legalmente autorizados existentes en España, por orden alfabético."
TITULO_HOTELES_EN = "Hotels legally authorized existing in Spain, in alphabetical order."
TITULO_POB_ES = "Poblaciones de España con hoteles legalmente autorizados, por orden alfabético."
TITULO_POB_EN = "Spanish towns with legally authorized hotels, in alphabetical order."

# Títulos (sin línea separadora)
y_start_index = cabecera_indice(pdf, TITULO_HOTELES_ES, TITULO_HOTELES_EN)

# Lista de hoteles ordenada
hoteles_lista = sorted(hotel_pages.keys(), key=lambda x: x.lower())

# Configuración: columnas verticales
COLS_INDEX = COLS_INDICE
col_width_index = (CONTENT_WIDTH - (COLS_INDEX - 1) * SEP_INDICE) / COLS_INDEX
row_height_index = ROW_H_INDICE  # Ajustado para tipografía pequeña
y_limit_index = Y_LIMIT_INDICE


# ---- FUNCIÓN DE FORMATO (tipografía 6pt equivalente) ----
def format_index_entry(pdf, name, page, max_width):
    encoded_name = name.encode("latin-1", "ignore").decode("latin-1")
    page_str = str(page)

    # Reservar espacio para número de página
    space_reserved = pdf.get_string_width(page_str) + 1.0
    max_name_width = max_width - space_reserved - 1.5

    # Truncado si hace falta
    while pdf.get_string_width(encoded_name) > max_name_width:
        encoded_name = encoded_name[:-1].rstrip()
        if len(encoded_name) <= 2:
            break
    if pdf.get_string_width(encoded_name) > max_name_width:
        encoded_name = encoded_name[:-2] + ".."

    # Puntos
    space_left = (
        max_width
        - pdf.get_string_width(encoded_name)
        - pdf.get_string_width(page_str)
        - 1
    )
    dot_count = max(2, int(space_left / pdf.get_string_width(".")))

    return f"{encoded_name} {'.' * dot_count} {page_str}"


# ---- IMPRIMIR ÍNDICE EN COLUMNAS VERTICALES ----
pdf.set_font("Helvetica", "", FONT_INDICE)
pdf.set_text_color(0, 0, 0)

pdf.set_y(y_start_index)


def columnas_indice(page_no, n_cols, ancho_col):
    base = x_contenido(page_no)
    return [base + i * (ancho_col + SEP_INDICE) for i in range(n_cols)]


x_cols = columnas_indice(pdf.page_no(), COLS_INDEX, col_width_index)
y_cols = [y_start_index] * COLS_INDEX

hotel_idx = 0
current_col = 0
page_count = 0

while hotel_idx < len(hoteles_lista):

    # Verificar si necesita página nueva (cualquier columna sobrepasa límite)
    if y_cols[current_col] + row_height_index > y_limit_index and hotel_idx < len(
        hoteles_lista
    ):

        # Pasar a siguiente columna
        current_col += 1

        if current_col >= COLS_INDEX:
            # Nueva página
            pdf.add_page()
            y_nueva = cabecera_indice(pdf, TITULO_HOTELES_ES, TITULO_HOTELES_EN)
            pdf.set_font("Helvetica", "", FONT_INDICE)

            current_col = 0
            x_cols = columnas_indice(pdf.page_no(), COLS_INDEX, col_width_index)
            y_cols = [y_nueva] * COLS_INDEX
            page_count += 1

    # Imprimir hotel en columna actual
    hotel = hoteles_lista[hotel_idx]
    pagina = hotel_pages[hotel]
    linea = format_index_entry(pdf, hotel, pagina, col_width_index - 2)

    pdf.set_xy(x_cols[current_col], y_cols[current_col])
    pdf.cell(col_width_index, row_height_index, linea, border=0, align="L")

    y_cols[current_col] += row_height_index
    hotel_idx += 1

# --- PORTADA ÍNDICE ALFABÉTICO DE POBLACIONES (estilo minimalista) ---
pdf.provincia_actual = None
pdf.add_page()
dibujar_portada_seccion(
    pdf,
    PORTADA_POBLACIONES_ES,
    PORTADA_POBLACIONES_EN,
    pdf.page_no(),
)

# --- INICIAR ÍNDICE ALFABÉTICO DE POBLACIONES ---
pdf.provincia_actual = None
pdf.add_page()

# Títulos del índice de poblaciones
y_start_pob_inicial = cabecera_indice(pdf, TITULO_POB_ES, TITULO_POB_EN)

# Poblaciones → página REAL (capturada durante el render del catálogo).
# loc_pages usa la localidad tal cual aparece; normalizamos la clave para
# fusionar variantes por espacios/mayúsculas y quedarnos con la 1ª página.
poblacion_pages = {}
for _loc, _pg in loc_pages.items():
    _clave = str(_loc).strip()
    if _clave and _clave not in poblacion_pages:
        poblacion_pages[_clave] = _pg

# Lista de poblaciones ordenada alfabéticamente (sin tildes)
poblaciones_lista = sorted(poblacion_pages.keys(), key=lambda x: normalizar_ciudad(x))

# Configuración: columnas verticales (igual que el índice de hoteles)
COLS_POB = COLS_INDICE
col_width_pob = (CONTENT_WIDTH - (COLS_POB - 1) * SEP_INDICE) / COLS_POB
row_height_pob = ROW_H_INDICE
y_start_pob = y_start_pob_inicial
y_limit_pob = Y_LIMIT_INDICE

# ---- IMPRIMIR ÍNDICE DE POBLACIONES EN COLUMNAS VERTICALES ----
pdf.set_font("Helvetica", "", FONT_INDICE)
pdf.set_text_color(0, 0, 0)
pdf.set_y(y_start_pob)

x_cols_pob = columnas_indice(pdf.page_no(), COLS_POB, col_width_pob)
y_cols_pob = [y_start_pob] * COLS_POB

pob_idx = 0
current_col_pob = 0

while pob_idx < len(poblaciones_lista):

    if y_cols_pob[current_col_pob] + row_height_pob > y_limit_pob and pob_idx < len(
        poblaciones_lista
    ):
        current_col_pob += 1

        if current_col_pob >= COLS_POB:
            pdf.add_page()
            y_nueva_pob = cabecera_indice(pdf, TITULO_POB_ES, TITULO_POB_EN)
            pdf.set_font("Helvetica", "", FONT_INDICE)

            current_col_pob = 0
            x_cols_pob = columnas_indice(pdf.page_no(), COLS_POB, col_width_pob)
            y_cols_pob = [y_nueva_pob] * COLS_POB

    poblacion = poblaciones_lista[pob_idx]
    pagina_pob = poblacion_pages[poblacion]
    linea_pob = format_index_entry(pdf, poblacion, pagina_pob, col_width_pob - 2)

    pdf.set_xy(x_cols_pob[current_col_pob], y_cols_pob[current_col_pob])
    pdf.cell(col_width_pob, row_height_pob, linea_pob, border=0, align="L")

    y_cols_pob[current_col_pob] += row_height_pob
    pob_idx += 1

pdf.output(PDF_FILE)
print("PDF generado con índice alfabético de 5 columnas verticales:", PDF_FILE)