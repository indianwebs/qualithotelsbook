import math
import os
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

# ---------------------------------------------------------------------------
# CORRECCIONES MANUALES
# ---------------------------------------------------------------------------
# Retoques puntuales sobre los datos. Se aplican aqui y NO en el excel de
# origen, para que no se pierdan si algun dia se vuelve a exportar la hoja.
# Clave = ID del hotel; valor = {columna: nuevo valor}.
CORRECCIONES = {
    2677: {"DIRECCION": "PADILLA, 173"},   # HOTEL GLORIES (Barcelona)
}

for _id, _campos in CORRECCIONES.items():
    _fila = df["ID"] == _id
    if not _fila.any():
        print(f"AVISO: no hay ningun hotel con ID {_id}; correccion ignorada")
        continue
    for _col, _valor in _campos.items():
        df.loc[_fila, _col] = _valor


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


# ---------------------------------------------------------------------------
# FUENTE INCRUSTADA
# ---------------------------------------------------------------------------
# fpdf2 usa por defecto Helvetica, que es una de las 14 fuentes "base" del
# estándar PDF: se nombra en el archivo pero NO se incrusta. KDP exige que
# todas las fuentes vayan incrustadas y, si no lo están, mete una sustituta
# por su cuenta y avisa de que "puede haber causado pequeños cambios".
#
# Usamos Arial, que es un clon métrico de Helvetica: comprobado sobre las
# líneas reales del libro, el ancho de cada cadena es IDÉNTICO al bit en los
# cinco cuerpos que usa el catálogo, así que la paginación no se mueve.
# Su fsType es 8 (Editable Embedding), o sea que incrustarla está permitido.
#
# Si las fuentes no estuvieran disponibles (otro equipo, Linux...), se vuelve
# a la Helvetica base: el PDF se genera igual, pero KDP volverá a avisar.
_RUTAS_FUENTE = {
    "": "C:/Windows/Fonts/arial.ttf",
    "B": "C:/Windows/Fonts/arialbd.ttf",
    "I": "C:/Windows/Fonts/ariali.ttf",
}
INCRUSTAR_FUENTE = all(os.path.exists(_r) for _r in _RUTAS_FUENTE.values())
FUENTE = "Arial" if INCRUSTAR_FUENTE else "Helvetica"


# --- PDF ---
class PDF(FPDF):
    def __init__(self):
        # Tamaño de página nativo 6" x 9" (KDP paperback)
        super().__init__(orientation="P", unit="mm", format=(PAGE_WIDTH, PAGE_HEIGHT))
        if INCRUSTAR_FUENTE:
            for _estilo, _ruta in _RUTAS_FUENTE.items():
                self.add_font(FUENTE, _estilo, _ruta)
        self.set_margins(MARGIN_GUTTER, Y_TOP, MARGIN_OUTER + BLEED)

    def header(self):
        # Márgenes simétricos: el medianil cambia de lado en cada página
        izq, der = margenes_pagina(self.page_no())
        self.set_margins(izq, Y_TOP, der)
        self.set_xy(izq, Y_TOP)

        # fpdf2 dibuja el pie de la página N cuando ya se ha pedido la N+1, y
        # para entonces `provincia_actual` puede haber cambiado. Decidimos aquí
        # (al abrir la página) si esta página lleva número, y el pie lo consulta.
        con_cabecera = getattr(self, "provincia_actual", "") not in [None, "", False]
        # Los indices alfabeticos finales no llevan cabecera de provincia, pero
        # SI deben numerarse: sin numero, la guia aparentaba acabar en la 467.
        self._pie_visible = con_cabecera or getattr(self, "pie_forzado", False)

        if not con_cabecera:
            return

        # Encabezado por provincia
        self.set_font(FUENTE, "B", FONT_CABECERA)
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
        if not getattr(self, "_pie_visible", False):
            return
        # Las portadas azules llevan su propio número (blanco, dentro del panel)
        if self.page_no() in getattr(self, "paginas_sin_pie", set()):
            return
        izq = x_contenido(self.page_no())
        self.set_xy(izq, Y_PIE)
        self.set_font(FUENTE, "I", 7)
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
# TAMAÑO DE PÁGINA Y MÁRGENES (KDP paperback 6" x 9" SIN SANGRADO)
# ---------------------------------------------------------------------------
# Tamaño final del libro (trim): 6" x 9" = 152.4 x 228.6 mm.
#
# El PDF se genera SIN SANGRADO: la página mide EXACTAMENTE el tamaño de corte
# (152.4 x 228.6 mm = 6" x 9"), sin ningún milímetro de más. Al subirlo a KDP
# hay que elegir "Sin sangrado / No bleed" y tamaño de recorte 6 x 9 pulgadas.
# Así Amazon no reescala nada y la página sale tal cual está diseñada.
#
# Las portadas azules siguen SIN margen blanco porque el rectángulo azul se
# pinta cubriendo la página entera (de borde a borde) y además desbordando un
# poco por fuera; el visor recorta el sobrante.
#
# Márgenes (medidos desde el borde de la página) exigidos por KDP:
#   - Medianil (margen interior, junto al lomo): 0.75" = 19.05 mm (501-700 págs)
#   - Margen exterior / superior / inferior:     mínimo 0.25" = 6.35 mm
# Usamos 19.3 mm de medianil y 12.7 mm (0.5") en exterior/superior/inferior.
#
# El medianil alterna de lado: en páginas impares (derechas) el interior es el
# borde izquierdo; en páginas pares (izquierdas), el derecho.
#
# Si algún día se quisiera volver al modo con sangrado, basta con poner
# CON_SANGRADO = True: el resto de la geometría se recalcula sola.
# ---------------------------------------------------------------------------
TRIM_WIDTH = 152.4
TRIM_HEIGHT = 228.6

CON_SANGRADO = True     # False = PDF a tamaño de corte exacto (subir "sin sangrado")
BLEED = 3.175 if CON_SANGRADO else 0.0

# Tamaño físico del PDF (incluye el sangrado)
PAGE_WIDTH = TRIM_WIDTH + BLEED
PAGE_HEIGHT = TRIM_HEIGHT + 2 * BLEED

# Márgenes laterales: KDP impone el medianil según el número de páginas y con
# 501-700 exige 0.75" = 19.05 mm, así que 19.3 es el mínimo legal con holgura.
# El exterior se mantiene en 17.1 mm para que la mancha quede prácticamente
# centrada en la página (solo 2.2 mm de diferencia entre ambos lados).
MARGIN_GUTTER = 19.3    # medianil (interior, junto al lomo)
MARGIN_OUTER = 17.1     # margen exterior
# Superior e inferior al mínimo práctico para ganar altura de columna: KDP
# exige 0.25" (6.35 mm) sin sangrado, así que 8 / 7 mm dejan holgura suficiente
# frente a la tolerancia de la guillotina sin desperdiciar página.
MARGIN_TOP = 8.0
MARGIN_BOTTOM = 7.0

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
# El pie tiene su PROPIO margen, independiente de MARGIN_BOTTOM. Si se bajara
# tocando MARGIN_BOTTOM, el catálogo ganaría altura, entrarían más hoteles por
# columna y cambiaría la paginación (y con ella el lomo de la cubierta). Así el
# número baja solo él.
ALTO_PIE = 4.5          # alto de la celda que contiene el número
MARGIN_PIE = 5.6        # del borde de corte inferior a la base de esa celda
Y_PIE = PAGE_HEIGHT - BLEED - MARGIN_PIE - ALTO_PIE

# El límite del texto sigue atado a MARGIN_BOTTOM, como hasta ahora.
Y_LIMIT = Y_BOTTOM - ALTO_PIE - 1.0

# Tipografías del catálogo (ajustadas al ancho real de columna de 6"x9"
# y a la densidad necesaria para mantener el libro por debajo de 600 páginas)
FONT_CABECERA = 9.5
FONT_LOCALIDAD = 6.7
FONT_NOMBRE = 6.1
FONT_CAT = 5.5
FONT_DETALLE = 5.5

line_height = 2.65
# Separación vertical entre el final de un hotel y el comienzo del siguiente.
# Bajarla aprovecha el hueco que antes quedaba muerto al pie de cada columna.
SEP_HOTELES = 1.2

# Justificación vertical: cuando una columna se cierra por estar llena, el
# hueco que sobra al pie se reparte a partes iguales entre los hoteles, de modo
# que todas las columnas llenas terminan a la misma altura. `SEP_EXTRA_MAX`
# limita cuánto puede crecer cada separación, por si alguna columna se cierra
# con un hueco anormalmente grande.
JUSTIFICAR_COLUMNAS = True
SEP_EXTRA_MAX = 6.0
# Pequeño colchón para que ninguna línea toque el borde de la mancha
ancho_texto = COLUMN_WIDTH - 1.5

# ---------------------------------------------------------------------------
# PALETA DE COLOR
# ---------------------------------------------------------------------------
# "color" -> azul de las fotos. Obliga a elegir tinta de color en KDP, que en
#            un libro de 524 paginas dispara el coste de impresion.
# "bn"    -> pensado para imprimir en blanco y negro. NO dejamos que Amazon
#            convierta el azul por su cuenta: ese azul se volveria un gris 130,
#            MAS CLARO que el texto negro (gris 30), y las cabeceras y
#            localidades quedarian mas desvaidas que los datos que encabezan.
#            Aqui van en negro puro, que destaca por peso, y las portadas de
#            seccion en gris oscuro con el texto en blanco.
IMPRESION = "color"             # "color" o "bn"

if IMPRESION == "bn":
    AZUL_PORTADA = (55, 55, 55)     # fondo de las portadas de seccion
    AZUL_ACENTO = (0, 0, 0)         # cabeceras de provincia y localidades
else:
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
def _dibujar_separador(pdf, x0, ancho, y, color=(255, 255, 255), escala=1.0):
    """Separador decorativo: línea — rombo — línea, centrado en `y`.

    Por defecto va en blanco y a tamaño completo (portadas azules de sección).
    `color` y `escala` permiten reutilizarlo en el índice general, donde va en
    azul y algo más pequeño para no competir con el título.
    """
    pdf.set_draw_color(*color)
    pdf.set_fill_color(*color)
    pdf.set_line_width(max(0.2, 0.5 * escala))
    cx = x0 + ancho / 2
    # Factores calculados sobre el ANCHO DE CONTENIDO (116 mm) para que el
    # separador conserve la misma longitud visual (~100 mm) que tenía cuando se
    # dibujaba sobre el ancho del recorte.
    largo = ancho * 0.368 * escala  # longitud de cada media línea
    hueco = ancho * 0.066 * escala  # separación entre línea y rombo
    pdf.line(cx - hueco - largo, y, cx - hueco, y)
    pdf.line(cx + hueco, y, cx + hueco + largo, y)
    # Rombo central (4 vértices)
    s = 2.0 * escala
    pdf.polygon(
        [(cx, y - s), (cx + s, y), (cx, y + s), (cx - s, y)],
        style="F",
    )


def _tamano_fuente_ajustado(pdf, lineas, ancho_util, size_max=13, size_min=6):
    """Mayor tamaño de fuente (Helvetica Bold) con el que TODAS las líneas
    caben en `ancho_util`."""
    size = size_max
    while size > size_min:
        pdf.set_font(FUENTE, "B", size)
        if all(pdf.get_string_width(_enc(l)) <= ancho_util for l in lineas):
            return size
        size -= 0.5
    return size_min


def dibujar_portada_seccion(pdf, lineas_es, lineas_en, page_number_display):
    """Portada azul de sección, bilingüe, al estilo de las fotos:
    - Azul de borde a borde: cubre TODA la página, sin ningún margen blanco
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

    # Fondo azul cubriendo la PÁGINA ENTERA, de borde a borde: cero margen
    # blanco. Se pinta además 2 mm MÁS GRANDE que el papel por los cuatro
    # lados; el visor recorta el sobrante, y así nunca queda una línea blanca
    # de un píxel en el borde por redondeo del RIP de la imprenta.
    pdf.set_fill_color(*AZUL_PORTADA)
    pdf.rect(-2, -2, PAGE_WIDTH + 4, PAGE_HEIGHT + 4, "F")
    pdf.set_text_color(255, 255, 255)

    # El texto se centra respecto al ÁREA DE CORTE (lo que queda del papel tras
    # guillotinar), no respecto a la mancha: así se ve perfectamente centrado.
    # El ancho útil se limita al medianil por ambos lados, de modo que el
    # bloque centrado sigue respetando el margen interior.
    # Los bloques se centran sobre el ÁREA DE CONTENIDO, que es el eje que usa
    # todo el resto del libro (cabeceras, títulos, tablas y números de página),
    # NO sobre el centro del papel. Como el medianil es mayor que el margen
    # exterior, ambos ejes se separan 1.1 mm; usar el del contenido es lo
    # correcto porque el lomo se traga parte del medianil al encuadernar.
    x0 = x_contenido(pdf.page_no())
    y0 = BLEED
    ancho = CONTENT_WIDTH
    alto = TRIM_HEIGHT

    # Número de página arriba a la derecha, a ras del borde del contenido,
    # exactamente en la misma posición que en la página del índice 1.
    pdf.set_font(FUENTE, "", 9)
    pdf.set_xy(x0 + CONTENT_WIDTH - 15, Y_TOP)
    pdf.cell(15, 6, str(page_number_display), align="R")

    ancho_util = ancho - 2
    size = _tamano_fuente_ajustado(pdf, lineas_es + lineas_en, ancho_util)
    alt = size * 0.62  # alto de línea proporcional al cuerpo

    # Bloque español (centrado alrededor del 30% de la mancha)
    pdf.set_font(FUENTE, "B", size)
    y_es = y0 + alto * 0.30 - (len(lineas_es) * alt) / 2
    pdf.set_xy(x0, y_es)
    for linea in lineas_es:
        pdf.set_x(x0)
        pdf.cell(ancho, alt, _enc(linea), align="C", new_x="LEFT", new_y="NEXT")

    # Separador decorativo en el centro vertical de la mancha
    _dibujar_separador(pdf, x0, ancho, y0 + alto * 0.505)

    # Bloque inglés (centrado alrededor del 68% de la mancha)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(FUENTE, "B", size)
    y_en = y0 + alto * 0.68 - (len(lineas_en) * alt) / 2
    pdf.set_xy(x0, y_en)
    for linea in lineas_en:
        pdf.set_x(x0)
        pdf.cell(ancho, alt, _enc(linea), align="C", new_x="LEFT", new_y="NEXT")

    # Resetear estilo
    pdf.set_text_color(0, 0, 0)
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.2)



# ---------------------------------------------------------------------------
# COLOCACIÓN DE LAS PORTADAS AZULES: SIEMPRE EN PÁGINA IMPAR (ANVERSO)
# ---------------------------------------------------------------------------
# En un libro encuadernado las páginas IMPARES son las de la derecha (recto /
# anverso) y las PARES las de la izquierda (verso / dorso). Para que la portada
# azul se vea siempre "de frente" al pasar la hoja y su dorso quede en blanco:
#   1. Si la portada fuese a caer en página par, se mete antes una hoja blanca.
#   2. Después de la portada se mete SIEMPRE otra hoja blanca (su reverso).
# Así el contenido siguiente vuelve a empezar en impar.
def pagina_en_blanco(pdf):
    """Añade una página totalmente vacía (sin cabecera, sin pie, sin número)."""
    pdf.provincia_actual = None
    pdf.provincia_continuacion = False
    pdf.pie_forzado = False
    pdf.add_page()


def nueva_portada_seccion(pdf, lineas_es, lineas_en):
    """Añade una portada azul garantizando que cae en página IMPAR (anverso)
    y que su reverso queda en blanco. Devuelve el número de esa portada, que es
    la página a la que apunta el índice general para esa sección."""
    pdf.provincia_actual = None
    pdf.provincia_continuacion = False
    pdf.pie_forzado = False   # la portada azul lleva su numero arriba a la dcha.
    # La portada sería la página siguiente a la actual: si fuese par, relleno.
    if (pdf.page_no() + 1) % 2 == 0:
        pagina_en_blanco(pdf)
    pdf.add_page()
    pagina_portada = pdf.page_no()
    dibujar_portada_seccion(pdf, lineas_es, lineas_en, pagina_portada)
    # Reverso en blanco: el contenido siguiente arrancará de nuevo en impar.
    pagina_en_blanco(pdf)
    return pagina_portada

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


# ---------------------------------------------------------------------------
# ÍNDICE GENERAL: DENOMINACIONES OFICIALES (ESPAÑOL / INGLÉS)
# ---------------------------------------------------------------------------
# Clave = provincia normalizada, sin tildes ni espacios (igual que CAPITALES).
# Cada valor es (nombre, nombre_en, capital, tipo):
#   nombre     nombre de la provincia en español
#   nombre_en  su forma inglesa; "" cuando no cambia
#   capital    capital, con la variante entre paréntesis cuando la tiene
#   tipo       ""       provincia
#              "uni"    Comunidad Autónoma uniprovincial
#              "ciudad" Ciudad Autónoma (Ceuta y Melilla)
INDICE_PROVINCIAS = {
    "ACORUNA":             ("A Coruña", "", "A Coruña", ""),
    "ALBACETE":            ("Albacete", "", "Albacete", ""),
    "ALICANTE":            ("Alicante", "", "Alicante", ""),
    "ALMERIA":             ("Almería", "", "Almería", ""),
    "ARABA":               ("Araba", "", "Vitoria (Gasteiz)", ""),
    "ASTURIAS":            ("Asturias", "", "Oviedo", "uni"),
    "AVILA":               ("Ávila", "", "Ávila", ""),
    "BADAJOZ":             ("Badajoz", "", "Badajoz", ""),
    "BARCELONA":           ("Barcelona", "", "Barcelona", ""),
    "BIZKAIA":             ("Bizkaia", "", "Bilbao", ""),
    "VIZCAYA":             ("Bizkaia", "", "Bilbao", ""),
    "BURGOS":              ("Burgos", "", "Burgos", ""),
    "CACERES":             ("Cáceres", "", "Cáceres", ""),
    "CADIZ":               ("Cádiz", "", "Cádiz", ""),
    "CANTABRIA":           ("Cantabria", "", "Santander", "uni"),
    "CASTELLON":           ("Castellón", "", "Castellón de la Plana", ""),
    "CEUTA":               ("Ceuta", "", "Ceuta", "ciudad"),
    "CIUDADREAL":          ("Ciudad Real", "", "Ciudad Real", ""),
    "CORDOBA":             ("Córdoba", "", "Córdoba", ""),
    "CUENCA":              ("Cuenca", "", "Cuenca", ""),
    "GIPUZKOA":            ("Gipuzkoa", "", "San Sebastián (Donostia)", ""),
    "GUIPUZCOA":           ("Gipuzkoa", "", "San Sebastián (Donostia)", ""),
    "GIRONA":              ("Girona", "", "Girona (Gerona)", ""),
    "GERONA":              ("Girona", "", "Girona (Gerona)", ""),
    "GRANADA":             ("Granada", "", "Granada", ""),
    "GUADALAJARA":         ("Guadalajara", "", "Guadalajara", ""),
    "HUELVA":              ("Huelva", "", "Huelva", ""),
    "HUESCA":              ("Huesca", "", "Huesca", ""),
    "ISLASBALEARES":       ("Islas Baleares", "the Balearic Islands",
                            "Palma de Mallorca", "uni"),
    "JAEN":                ("Jaén", "", "Jaén", ""),
    "LARIOJA":             ("La Rioja", "", "Logroño", "uni"),
    "LASPALMAS":           ("Las Palmas", "", "Las Palmas de Gran Canaria", ""),
    "LEON":                ("León", "", "León", ""),
    "LLEIDA":              ("Lleida", "", "Lleida (Lérida)", ""),
    "LUGO":                ("Lugo", "", "Lugo", ""),
    "MADRID":              ("Madrid", "", "Madrid", "uni"),
    "MALAGA":              ("Málaga", "", "Málaga", ""),
    "MELILLA":             ("Melilla", "", "Melilla", "ciudad"),
    "MURCIA":              ("Murcia", "", "Murcia", "uni"),
    "NAVARRA":             ("Navarra", "", "Pamplona (Iruña)", "uni"),
    "OURENSE":             ("Ourense", "", "Ourense (Orense)", ""),
    "PALENCIA":            ("Palencia", "", "Palencia", ""),
    "PONTEVEDRA":          ("Pontevedra", "", "Pontevedra", ""),
    "SALAMANCA":           ("Salamanca", "", "Salamanca", ""),
    "SANTACRUZDETENERIFE": ("Santa Cruz de Tenerife", "",
                            "Santa Cruz de Tenerife", ""),
    "SEGOVIA":             ("Segovia", "", "Segovia", ""),
    "SEVILLA":             ("Sevilla", "", "Sevilla", ""),
    "SORIA":               ("Soria", "", "Soria", ""),
    "TARRAGONA":           ("Tarragona", "", "Tarragona", ""),
    "TERUEL":              ("Teruel", "", "Teruel", ""),
    "TOLEDO":              ("Toledo", "", "Toledo", ""),
    "VALENCIA":            ("Valencia", "", "Valencia", ""),
    "VALLADOLID":          ("Valladolid", "", "Valladolid", ""),
    "ZAMORA":              ("Zamora", "", "Zamora", ""),
    "ZARAGOZA":            ("Zaragoza", "", "Zaragoza", ""),
}

# Las dos secciones finales también figuran en el índice, tras un hueco.
ENTRADAS_FINALES = [
    {
        "clave": "hoteles",
        "es": "Nombre de los hoteles de España, por orden alfabético",
        "en": "Names of hotels in Spain, in alphabetical order",
        "pagina": None,
    },
    {
        "clave": "poblaciones",
        "es": "Nombre de las ciudades y localidades de España, por orden alfabético",
        "en": "Names of cities and towns in Spain, in alphabetical order",
        "pagina": None,
    },
]


def entrada_indice(prov):
    """Línea del índice general (ES e EN) para una provincia del catálogo."""
    clave = normalizar_provincia(prov).replace(" ", "")
    nombre, nombre_en, capital, tipo = INDICE_PROVINCIAS.get(
        clave, (corregir_preposiciones(prov), "", CAPITALES.get(clave, ""), "")
    )
    nombre_en = nombre_en or nombre
    if tipo == "ciudad":
        es = f"Ciudad Autónoma de {nombre}"
        en = f"Autonomous City of {nombre_en}"
    else:
        es = f"Provincia de {nombre}"
        en = f"Province of {nombre_en}"
        if tipo == "uni":
            es += ", Comunidad Autónoma uniprovincial"
            en += ", single-province Autonomous Community"
        if capital:
            es += f", capital {capital}"
            en += f", capital {capital}"
    return {"provincia": prov, "es": es, "en": en, "pagina": None}


# Obtener lista única de provincias en orden alfabético (sin tildes)
provincias_unicas = sorted(df["PROVINCIA"].unique().tolist(), key=normalizar_provincia)

# Entradas del índice general: una por provincia, hueco, y las dos secciones
# finales. Los diccionarios se rellenan con la página REAL tras la pasada 1;
# como se reutilizan los mismos objetos, basta con mutarlos.
indice_provincias = [entrada_indice(prov) for prov in provincias_unicas]
entradas_indice = indice_provincias + [None] + ENTRADAS_FINALES

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


def alto_lineas(pdf, texto, ancho):
    """Nº de líneas que ocupará `texto` al dibujarlo, sin dibujarlo."""
    if not texto:
        return 0
    return len(pdf.multi_cell(ancho, line_height, texto,
                              dry_run=True, output="LINES"))


def alto_real_hotel(pdf, d):
    """Alto EXACTO (mm) del bloque de un hotel. A diferencia de
    `calcular_altura_bloque`, que estima por lo alto para decidir saltos, esto
    mide lo que de verdad va a ocupar: es lo que permite saber cuánto hueco
    sobra al pie de la columna y repartirlo."""
    n = 0
    if d["cat"]:
        pdf.set_font(FUENTE, "B", FONT_CAT)
        n += alto_lineas(pdf, d["cat"], ancho_texto)
    pdf.set_font(FUENTE, "B", FONT_NOMBRE)
    n += alto_lineas(pdf, d["nombre"], ancho_texto)
    pdf.set_font(FUENTE, "", FONT_DETALLE)
    for clave in ("reg", "dir", "loc", "tel", "web"):
        n += alto_lineas(pdf, d[clave], ancho_texto)
    return n * line_height


def alto_real_localidad(pdf, localidad):
    """Alto EXACTO (mm) del rótulo de localidad, incluido el 1 mm de aire."""
    pdf.set_font(FUENTE, "B", FONT_LOCALIDAD)
    return 1 + alto_lineas(pdf, _enc(localidad.upper()), COLUMN_WIDTH) * line_height


def render_catalogo(pdf):
    """Dibuja TODO el catálogo por provincias en `pdf`.

    Devuelve (prov_pages, hotel_pages, loc_pages): la página REAL de la primera
    aparición de cada provincia, hotel (nombre limpio) y localidad.

    Los hoteles no se dibujan según van saliendo: se acumulan en un buffer por
    columna y se vuelcan cuando la columna se cierra. Así, en el momento de
    dibujar, se sabe cuánto hueco sobra al pie y se reparte a partes iguales
    entre los hoteles (justificación vertical), de modo que todas las columnas
    llenas terminan a la misma altura. El reparto NO altera la paginación: las
    decisiones de salto de columna y de página son exactamente las de antes.
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

    # --- buffer de justificación vertical ---
    buffer = [[] for _ in range(COLS)]
    y_inicio = [Y_START] * COLS

    def dibujar_item(x, y, item):
        """Dibuja un rótulo de localidad (si lo lleva) y el bloque del hotel."""
        if item["loc"] is not None:
            y += 1
            pdf.set_xy(x, y)
            pdf.set_font(FUENTE, "B", FONT_LOCALIDAD)
            pdf.set_text_color(*AZUL_ACENTO)
            pdf.multi_cell(COLUMN_WIDTH, line_height, _enc(item["loc"]),
                           border=0, align="L")
            y = pdf.get_y()

        d = item["d"]
        pdf.set_xy(x, y)
        pdf.set_text_color(0, 0, 0)
        if d["cat"]:
            pdf.set_font(FUENTE, "B", FONT_CAT)
            pdf.multi_cell(ancho_texto, line_height, d["cat"], border=0, align="L")
        pdf.set_x(x)
        pdf.set_font(FUENTE, "B", FONT_NOMBRE)
        pdf.multi_cell(ancho_texto, line_height, d["nombre"], border=0, align="L")
        pdf.set_font(FUENTE, "", FONT_DETALLE)
        for clave in ("reg", "dir", "loc", "tel", "web"):
            if d[clave]:
                pdf.set_x(x)
                pdf.multi_cell(ancho_texto, line_height, d[clave], border=0, align="L")
        return pdf.get_y()

    def volcar_columna(col, completa):
        """Vuelca a la página lo acumulado en la columna.

        `completa` indica que la columna se cerró porque no cabía nada más; solo
        en ese caso se reparte el hueco sobrante. La última columna de cada
        provincia se queda con la separación normal: repartir tres hoteles a lo
        largo de toda la página quedaría ridículo.
        """
        items = buffer[col]
        buffer[col] = []
        if not items:
            return
        sep = SEP_HOTELES
        huecos = len(items) - 1
        if completa and JUSTIFICAR_COLUMNAS and huecos > 0:
            usado = sum(it["alto"] for it in items) + huecos * SEP_HOTELES
            sobra = Y_LIMIT - y_inicio[col] - usado
            if sobra > 0:
                sep += min(sobra / huecos, SEP_EXTRA_MAX)
        y = y_inicio[col]
        x = x_positions[col]
        for item in items:
            y = dibujar_item(x, y, item) + sep

    for idx, row in df.iterrows():
        provincia = str(row["PROVINCIA"])
        localidad = str(row["LOCALIDAD"])
        hotel_name = str(row["NOMBRE DE EMPRESA"]).strip()

        # CAMBIO DE PROVINCIA → NUEVA PÁGINA Y RESET DE ALTURAS
        if provincia != provincia_anterior:
            # La columna a medias de la provincia anterior se vuelca ANTES de
            # cambiar de página, o su contenido caería en la página siguiente.
            volcar_columna(current_col, completa=False)
            provincia_anterior = provincia
            localidad_anterior = ""
            pdf.provincia_actual = provincia
            pdf.provincia_continuacion = False
            pdf.add_page()
            x_positions = columnas(pdf.page_no())
            current_col = 0
            y_actual = [Y_START] * COLS
            y_inicio = [Y_START] * COLS
            if provincia not in prov_pages:
                prov_pages[provincia] = pdf.page_no()

        hotel_name_display = limpiar_nombre_hotel(hotel_name)

        _d = construir_lineas_hotel(row)
        lineas_hotel = [
            _d["nombre"], _d["cat"], _d["reg"],
            _d["dir"], _d["loc"], _d["tel"], _d["web"],
        ]

        # Altura estimada del hotel (solo para decidir salto de columna/página)
        pdf.set_font(FUENTE, "", FONT_NOMBRE)
        altura_hotel = calcular_altura_bloque(
            pdf, [_l for _l in lineas_hotel if _l], ancho_texto, line_height
        )

        hay_cambio_localidad = localidad != localidad_anterior
        altura_localidad = 0
        if hay_cambio_localidad:
            altura_localidad = (
                calcular_altura_linea(pdf, localidad.upper(), COLUMN_WIDTH, line_height) + 4
            )

        altura_total_requerida = altura_localidad + altura_hotel + SEP_HOTELES
        localidad_cont = False

        if y_actual[current_col] + altura_total_requerida > Y_LIMIT:
            # La columna se cierra por estar llena: aquí sí se justifica.
            volcar_columna(current_col, completa=True)
            current_col += 1
            if current_col >= COLS:
                pdf.provincia_continuacion = True
                if not hay_cambio_localidad:
                    localidad_cont = True
                pdf.add_page()
                x_positions = columnas(pdf.page_no())
                current_col = 0
                y_actual = [Y_START] * COLS
                y_inicio = [Y_START] * COLS

        x = x_positions[current_col]

        # ---- REGISTRAR HOTEL CON SU PÁGINA REAL (ya resuelto el salto de página) ----
        if hotel_name_display and hotel_name_display not in hotel_pages:
            hotel_pages[hotel_name_display] = pdf.page_no()

        if hay_cambio_localidad:
            localidad_anterior = localidad
            if localidad not in loc_pages:
                loc_pages[localidad] = pdf.page_no()
        elif localidad_cont:
            # El "(cont.)" encabeza la página entera, no una columna: se dibuja
            # al vuelo y empuja hacia abajo el arranque de las tres columnas.
            pdf.set_xy(x_positions[0], Y_START + 1)
            pdf.set_font(FUENTE, "B", FONT_LOCALIDAD)
            pdf.set_text_color(*AZUL_ACENTO)
            pdf.multi_cell(COLUMN_WIDTH, line_height,
                           _enc(localidad.upper() + " (cont.)"), border=0, align="L")
            cont_y = pdf.get_y()
            for _c in range(COLS):
                y_actual[_c] = cont_y
                y_inicio[_c] = cont_y

        item = {
            "loc": localidad.upper() if hay_cambio_localidad else None,
            "d": _d,
        }
        item["alto"] = alto_real_hotel(pdf, _d)
        if hay_cambio_localidad:
            item["alto"] += alto_real_localidad(pdf, localidad)
        buffer[current_col].append(item)
        y_actual[current_col] += item["alto"] + SEP_HOTELES

    # La última columna del catálogo queda a medias: sin justificar.
    volcar_columna(current_col, completa=False)

    return prov_pages, hotel_pages, loc_pages


# ---------------------------------------------------------------------------
# ÍNDICES ALFABÉTICOS FINALES (hoteles y poblaciones)
# ---------------------------------------------------------------------------
# Van muy compactos (4 columnas) para no inflar el número total de páginas.
# Se definen como funciones porque los usan LAS DOS pasadas: la de medición
# necesita saber en qué página empieza cada sección para poder anunciarla en el
# índice general, y la definitiva las dibuja igual.
FONT_TITULO_INDICE = 7.5
FONT_INDICE = 5.0
ROW_H_INDICE = 2.9
COLS_INDICE = 4
SEP_INDICE = 2.5
Y_LIMIT_INDICE = Y_LIMIT

TITULO_HOTELES_ES = "Hoteles legalmente autorizados existentes en España, por orden alfabético."
TITULO_HOTELES_EN = "Hotels legally authorized existing in Spain, in alphabetical order."
TITULO_POB_ES = "Poblaciones de España con hoteles legalmente autorizados, por orden alfabético."
TITULO_POB_EN = "Spanish towns with legally authorized hotels, in alphabetical order."


def cabecera_indice(pdf, titulo_es, titulo_en):
    """Imprime los dos títulos bilingües y deja el cursor bajo ellos."""
    x = x_contenido(pdf.page_no())
    pdf.set_xy(x, Y_TOP)
    pdf.set_font(FUENTE, "B", FONT_TITULO_INDICE)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(CONTENT_WIDTH, 4.5, _enc(titulo_es), new_x="LEFT", new_y="NEXT", align="C")
    pdf.cell(CONTENT_WIDTH, 4.5, _enc(titulo_en), new_x="LEFT", new_y="NEXT", align="C")
    pdf.ln(1.5)
    return pdf.get_y()


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


def columnas_indice(page_no, n_cols, ancho_col):
    base = x_contenido(page_no)
    return [base + i * (ancho_col + SEP_INDICE) for i in range(n_cols)]


def render_indice_alfabetico(pdf, titulo_es, titulo_en, entradas):
    """Índice alfabético a 4 columnas verticales. `entradas` = [(texto, página)].

    Devuelve la página donde arranca."""
    pdf.provincia_actual = None
    pdf.pie_forzado = True          # a partir de aqui las paginas van numeradas
    pdf.add_page()
    primera_pagina = pdf.page_no()

    y_inicio = cabecera_indice(pdf, titulo_es, titulo_en)
    ancho_col = (CONTENT_WIDTH - (COLS_INDICE - 1) * SEP_INDICE) / COLS_INDICE

    pdf.set_font(FUENTE, "", FONT_INDICE)
    pdf.set_text_color(0, 0, 0)

    x_cols = columnas_indice(pdf.page_no(), COLS_INDICE, ancho_col)
    y_cols = [y_inicio] * COLS_INDICE
    col = 0

    for texto, pagina in entradas:
        # ¿Cabe otra línea en esta columna? Si no, se pasa a la siguiente y,
        # agotadas las cuatro, a una página nueva con su cabecera.
        if y_cols[col] + ROW_H_INDICE > Y_LIMIT_INDICE:
            col += 1
            if col >= COLS_INDICE:
                pdf.add_page()
                y_nueva = cabecera_indice(pdf, titulo_es, titulo_en)
                pdf.set_font(FUENTE, "", FONT_INDICE)
                col = 0
                x_cols = columnas_indice(pdf.page_no(), COLS_INDICE, ancho_col)
                y_cols = [y_nueva] * COLS_INDICE

        linea = format_index_entry(pdf, texto, pagina, ancho_col - 2)
        pdf.set_xy(x_cols[col], y_cols[col])
        pdf.cell(ancho_col, ROW_H_INDICE, linea, border=0, align="L")
        y_cols[col] += ROW_H_INDICE

    return primera_pagina


def render_secciones_finales(pdf, hotel_pages, loc_pages):
    """Portada azul + índice alfabético, para hoteles y para poblaciones.

    Devuelve (página de la portada de hoteles, página de la de poblaciones):
    son las que anuncia el índice general."""
    hoteles_lista = sorted(hotel_pages.keys(), key=lambda x: x.lower())
    entradas_hoteles = [(h, hotel_pages[h]) for h in hoteles_lista]

    # Poblaciones → página REAL (capturada durante el render del catálogo).
    # loc_pages usa la localidad tal cual aparece; normalizamos la clave para
    # fusionar variantes por espacios/mayúsculas y quedarnos con la 1ª página.
    poblacion_pages = {}
    for _loc, _pg in loc_pages.items():
        _clave = str(_loc).strip()
        if _clave and _clave not in poblacion_pages:
            poblacion_pages[_clave] = _pg
    poblaciones_lista = sorted(poblacion_pages.keys(), key=normalizar_ciudad)
    entradas_pob = [(p, poblacion_pages[p]) for p in poblaciones_lista]

    pag_hoteles = nueva_portada_seccion(pdf, PORTADA_HOTELES_ES, PORTADA_HOTELES_EN)
    render_indice_alfabetico(pdf, TITULO_HOTELES_ES, TITULO_HOTELES_EN, entradas_hoteles)

    pag_poblaciones = nueva_portada_seccion(pdf, PORTADA_POBLACIONES_ES, PORTADA_POBLACIONES_EN)
    render_indice_alfabetico(pdf, TITULO_POB_ES, TITULO_POB_EN, entradas_pob)

    return pag_hoteles, pag_poblaciones


# ---------------------------------------------------------------------------
# ÍNDICE GENERAL DE PROVINCIAS (primeras páginas del libro)
# ---------------------------------------------------------------------------
# Una línea por provincia, al estilo clásico de guía:
#
#   Provincia de Cantabria, Comunidad Autónoma uniprovincial, capital Santander
#                                                            ..... pág. 135
#
# Primero la página en español y después la misma en inglés. El cuerpo de letra
# se calcula solo: se busca el mayor con el que la entrada MÁS LARGA de los dos
# idiomas sigue cabiendo en una línea, así ninguna se parte ni se recorta. El
# alto de línea reparte las entradas por toda la mancha, para que la página
# quede llena y equilibrada.
TITULO_INDICE = {"es": "ÍNDICE", "en": "INDEX"}
SUBTITULO_INDICE = {
    "es": "Provincias y ciudades autónomas de España, sus capitales y su página",
    "en": "Provinces and autonomous cities of Spain, their capitals and their page",
}
ETIQUETA_PAG = {"es": "pág.", "en": "page"}

FONT_IDX_TITULO = 16
FONT_IDX_SUBTITULO = 7.0
FONT_IDX_MAX = 8.0      # cuerpo ideal de las entradas
FONT_IDX_MIN = 5.0      # cuerpo mínimo antes de rendirse
ROW_IDX_MAX = 4.8       # interlínea máxima (con pocas entradas no se desparrama)
ROW_IDX_MIN = 3.0       # interlínea mínima antes de pasar a otra página
ANCHO_MIN_PUNTOS = 6.0  # hueco mínimo reservado a los puntos guía


def _metricas_indice_general(pdf, textos):
    """Mayor cuerpo con el que TODAS las entradas caben en una sola línea.

    Devuelve (cuerpo, ancho de la etiqueta 'pág.', ancho de los dígitos). Los
    dos anchos son fijos, así que el número de página se alinea en columna y el
    reparto de páginas no depende de cuántas cifras tenga cada número."""
    size = FONT_IDX_MAX
    while True:
        pdf.set_font(FUENTE, "", size)
        w_digitos = pdf.get_string_width("000") + 1.0
        w_etiqueta = max(pdf.get_string_width(_enc(e)) for e in ETIQUETA_PAG.values()) + 1.5
        disponible = CONTENT_WIDTH - w_etiqueta - w_digitos - ANCHO_MIN_PUNTOS
        cabe = all(pdf.get_string_width(_enc(t)) <= disponible for t in textos)
        if cabe or size <= FONT_IDX_MIN:
            return size, w_etiqueta, w_digitos
        size -= 0.25


def _folio_superior(pdf):
    """Número de página arriba, SIEMPRE en el borde exterior: a la derecha en
    las impares (las de la derecha del libro) y a la izquierda en las pares,
    para que nunca caiga del lado del lomo."""
    x = x_contenido(pdf.page_no())
    pdf.set_font(FUENTE, "", 9)
    pdf.set_text_color(0, 0, 0)
    if pdf.page_no() % 2 == 1:
        pdf.set_xy(x + CONTENT_WIDTH - 15, Y_TOP)
        pdf.cell(15, 6, str(pdf.page_no()), align="R")
    else:
        pdf.set_xy(x, Y_TOP)
        pdf.cell(15, 6, str(pdf.page_no()), align="L")


def _cabecera_indice_general(pdf, idioma):
    """Número de página, título, subtítulo y filete. Devuelve la Y de arranque."""
    x = x_contenido(pdf.page_no())
    _folio_superior(pdf)

    pdf.set_xy(x, Y_TOP + 6)
    pdf.set_font(FUENTE, "B", FONT_IDX_TITULO)
    pdf.set_text_color(*AZUL_ACENTO)
    pdf.cell(CONTENT_WIDTH, 9, _enc(TITULO_INDICE[idioma]),
             align="C", new_x="LEFT", new_y="NEXT")

    pdf.set_font(FUENTE, "I", FONT_IDX_SUBTITULO)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(CONTENT_WIDTH, 4.2, _enc(SUBTITULO_INDICE[idioma]),
             align="C", new_x="LEFT", new_y="NEXT")

    y_filete = pdf.get_y() + 2.6
    _dibujar_separador(pdf, x, CONTENT_WIDTH, y_filete, color=AZUL_ACENTO, escala=0.55)

    pdf.set_text_color(0, 0, 0)
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.2)
    return y_filete + 4.5


def _fila_indice_general(pdf, x, y, entrada, idioma, size, row_h, w_etiqueta, w_digitos):
    """Una línea: texto — puntos guía — 'pág.' — número (alineado a la derecha)."""
    if entrada is None:          # hueco antes de las dos secciones finales
        return
    texto = _enc(entrada[idioma])
    pdf.set_font(FUENTE, "", size)
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(x, y)
    w_texto = pdf.get_string_width(texto)
    pdf.cell(w_texto + 0.5, row_h, texto, align="L")

    x_etiqueta = x + CONTENT_WIDTH - w_etiqueta - w_digitos

    # Puntos guía, en gris para que no pesen más que el texto
    x_puntos = x + w_texto + 1.4
    ancho_puntos = x_etiqueta - 1.0 - x_puntos
    w_punto = pdf.get_string_width(".")
    if ancho_puntos > w_punto:
        pdf.set_text_color(150, 150, 150)
        pdf.set_xy(x_puntos, y)
        pdf.cell(ancho_puntos, row_h, "." * int(ancho_puntos / w_punto), align="L")

    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(x_etiqueta, y)
    pdf.cell(w_etiqueta, row_h, _enc(ETIQUETA_PAG[idioma]), align="R")
    pdf.set_xy(x + CONTENT_WIDTH - w_digitos, y)
    pdf.set_font(FUENTE, "B", size)
    pdf.cell(w_digitos, row_h,
             "" if entrada["pagina"] is None else str(entrada["pagina"]), align="R")


def _render_indice_idioma(pdf, entradas, idioma, size, w_etiqueta, w_digitos):
    """Dibuja el índice completo en un idioma. Devuelve las páginas que ocupa."""
    pdf.provincia_actual = None
    pdf.provincia_continuacion = False
    pdf.pie_forzado = False      # el número va arriba, no al pie
    pdf.add_page()
    y = _cabecera_indice_general(pdf, idioma)

    disponible = Y_LIMIT - y
    por_pagina = max(1, int(disponible // ROW_IDX_MIN))
    if len(entradas) <= por_pagina:
        # Cabe entero: se reparte por toda la mancha (sin pasarse de ROW_IDX_MAX)
        row_h = min(ROW_IDX_MAX, disponible / max(1, len(entradas)))
    else:
        row_h = ROW_IDX_MIN

    paginas = 1
    x = x_contenido(pdf.page_no())
    for entrada in entradas:
        if y + row_h > Y_LIMIT + 0.01:
            pdf.add_page()
            paginas += 1
            y = _cabecera_indice_general(pdf, idioma)
            x = x_contenido(pdf.page_no())
        _fila_indice_general(pdf, x, y, entrada, idioma, size, row_h,
                             w_etiqueta, w_digitos)
        y += row_h
    return paginas


# ---------------------------------------------------------------------------
# PÁGINA DEL MAPA POLÍTICO
# ---------------------------------------------------------------------------
MAPA_IMAGEN = "mapa final.jpg"

# Marcadores del formato JPEG, escritos por su valor para no depender de
# secuencias de escape dentro de literales de bytes.
MARCA = bytes((0xFF,))          # todo marcador empieza por aquí
SOI = bytes((0xFF, 0xD8))       # "Start Of Image": los dos primeros bytes
TITULO_MAPA_ES = "MAPA POLÍTICO DE ESPAÑA"
TITULO_MAPA_EN = "POLITICAL MAP OF SPAIN"
FONT_MAPA_TITULO_EN = 11.0


def _tamano_jpeg(ruta):
    """Ancho y alto en píxeles de un JPEG, leyendo solo su cabecera.

    Hace falta para centrar el mapa en la página ANTES de dibujarlo. Se lee a
    mano para no arrastrar una librería de imágenes por dos números: basta con
    recorrer los marcadores del fichero hasta el SOF, que es donde el formato
    guarda las dimensiones.
    """
    with open(ruta, "rb") as f:
        if f.read(2) != SOI:
            raise ValueError(f"{ruta} no es un JPEG")
        while True:
            byte = f.read(1)
            while byte and byte != MARCA:
                byte = f.read(1)
            marcador = f.read(1)
            while marcador == MARCA:          # relleno entre marcadores
                marcador = f.read(1)
            if not marcador:
                raise ValueError(f"no se encontró el tamaño en {ruta}")
            # Los SOF (0xC0-0xCF) llevan las dimensiones, menos DHT, JPG y DAC
            if 0xC0 <= marcador[0] <= 0xCF and marcador[0] not in (0xC4, 0xC8, 0xCC):
                f.read(3)                     # longitud (2) + precisión (1)
                alto = int.from_bytes(f.read(2), "big")
                ancho = int.from_bytes(f.read(2), "big")
                return ancho, alto
            longitud = int.from_bytes(f.read(2), "big")
            f.seek(longitud - 2, 1)


def render_pagina_mapa(pdf):
    """Página del mapa político, detrás de los dos índices.

    Es una sola página con los títulos en los dos idiomas: el mapa es el mismo
    para ambos, así que duplicarlo solo gastaría papel. No cuesta ninguna
    página de más frente a no ponerlo, porque el catálogo arranca siempre en
    impar y el script metería igualmente una hoja de relleno.
    """
    pdf.provincia_actual = None
    pdf.provincia_continuacion = False
    pdf.pie_forzado = False          # el número va arriba, como en el índice
    pdf.add_page()
    _folio_superior(pdf)

    ancho_px, alto_px = _tamano_jpeg(MAPA_IMAGEN)
    alto_img = CONTENT_WIDTH * alto_px / ancho_px
    alto_es, alto_en = 9.0, 6.0

    # Centrado en el hueco que queda bajo el folio
    arriba = Y_TOP + 6.0
    y = arriba + max(0.0, (Y_LIMIT - arriba - alto_es - alto_en - alto_img) / 2)
    x = x_contenido(pdf.page_no())

    # Mismo cuerpo, color y negrita que el título del índice, para que las
    # páginas de los preliminares se lean como un conjunto
    pdf.set_text_color(*AZUL_ACENTO)
    pdf.set_xy(x, y)
    pdf.set_font(FUENTE, "B", FONT_IDX_TITULO)
    pdf.cell(CONTENT_WIDTH, alto_es, _enc(TITULO_MAPA_ES), align="C")
    pdf.set_xy(x, y + alto_es)
    pdf.set_font(FUENTE, "B", FONT_MAPA_TITULO_EN)
    pdf.cell(CONTENT_WIDTH, alto_en, _enc(TITULO_MAPA_EN), align="C")
    pdf.set_text_color(0, 0, 0)

    pdf.image(MAPA_IMAGEN, x=x, y=y + alto_es + alto_en, w=CONTENT_WIDTH)
    return 1


def render_indice_general(pdf, entradas):
    """Índice general: página(s) en español, luego en inglés, y detrás la
    página del mapa político, común a los dos idiomas.

    Devuelve el número total de páginas que ha ocupado."""
    textos = [e[idi] for e in entradas if e for idi in ("es", "en")]
    size, w_etiqueta, w_digitos = _metricas_indice_general(pdf, textos)
    paginas = 0
    for idioma in ("es", "en"):
        paginas += _render_indice_idioma(pdf, entradas, idioma, size,
                                         w_etiqueta, w_digitos)
    paginas += render_pagina_mapa(pdf)
    return paginas


# ---- Cuántas páginas ocupa el índice general (se necesita ANTES de medir) ----
# El reparto no depende de los números de página (van en una caja de ancho fijo),
# así que basta con maquetarlo una vez en vacío y contar páginas.
_scratch_indice = PDF()
_scratch_indice.set_auto_page_break(auto=False)
_scratch_indice.provincia_actual = None
N_PAGINAS_INDICE = render_indice_general(_scratch_indice, entradas_indice)
del _scratch_indice

# ---- PASADA 1: render de medición (a un PDF temporal) ----
# Páginas fijas antes del catálogo: [portada opc.] + [intro opc.] + índice general.
paginas_fijas_antes = (
    (1 if SHOW_PORTADA else 0)
    + (1 if SHOW_SEGUNDA_PAGINA else 0)
    + N_PAGINAS_INDICE
)
# La portada azul del catálogo se fuerza a página IMPAR y lleva reverso blanco,
# igual que en la pasada 2; hay que contarlo aquí o los números del índice
# de provincias saldrían desplazados.
if (paginas_fijas_antes + 1) % 2 == 0:
    paginas_fijas_antes += 1        # hoja blanca de relleno antes de la portada
paginas_fijas_antes += 2            # portada azul + su reverso en blanco

_scratch = PDF()
_scratch.set_auto_page_break(auto=False)
_scratch.set_font(FUENTE, "", 9)
_scratch.provincia_actual = None  # sin cabecera/pie en las páginas fijas dummy
for _ in range(paginas_fijas_antes):
    _scratch.add_page()
prov_pages_real, _hotel_pages_m, _loc_pages_m = render_catalogo(_scratch)
# Las dos secciones finales también se miden aquí: el índice general anuncia en
# qué página empiezan, y eso solo se sabe tras dibujar el catálogo entero.
_pag_hoteles, _pag_poblaciones = render_secciones_finales(
    _scratch, _hotel_pages_m, _loc_pages_m
)
del _scratch

# Índice general con las páginas REALES
for item in indice_provincias:
    prov = item["provincia"]
    if prov in prov_pages_real:
        item["pagina"] = prov_pages_real[prov]
ENTRADAS_FINALES[0]["pagina"] = _pag_hoteles
ENTRADAS_FINALES[1]["pagina"] = _pag_poblaciones

# ---- PASADA 2: generar el PDF completo en orden correcto ----

# --- CREAR PDF FINAL ---
pdf = PDF()
pdf.set_auto_page_break(auto=False)
pdf.set_font(FUENTE, "", 9)
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

# --- ÍNDICE GENERAL (español y después inglés) ---
_paginas_indice = render_indice_general(pdf, entradas_indice)
if _paginas_indice != N_PAGINAS_INDICE:
    print(f"AVISO: el índice ocupa {_paginas_indice} páginas y se habían "
          f"reservado {N_PAGINAS_INDICE}; los números estarán desplazados")

# --- PORTADA AZUL DEL CATÁLOGO (antes de las provincias) ---
nueva_portada_seccion(pdf, PORTADA_CATALOGO_ES, PORTADA_CATALOGO_EN)

# --- GENERAR CATÁLOGO (pasada 2, render final; páginas idénticas a la pasada 1) ---
prov_pages_final, hotel_pages, loc_pages = render_catalogo(pdf)

# --- SECCIONES FINALES: índices alfabéticos de hoteles y de poblaciones ---
pag_hoteles, pag_poblaciones = render_secciones_finales(pdf, hotel_pages, loc_pages)

# --- COMPROBACIÓN: las dos pasadas tienen que coincidir página a página ---
_desajustes = [p for p, n in prov_pages_final.items() if prov_pages_real.get(p) != n]
if _desajustes or (pag_hoteles, pag_poblaciones) != (_pag_hoteles, _pag_poblaciones):
    print(f"AVISO: el índice no cuadra con el catálogo ({len(_desajustes)} provincias "
          f"desplazadas). Revisar la paginación.")

# --- CIERRE: el interior debe tener un número PAR de páginas ---
# Cada hoja física lleva dos páginas; si el total fuese impar, la imprenta
# añadiría una hoja por su cuenta. Mejor añadirla nosotros, limpia.
if pdf.page_no() % 2 == 1:
    pagina_en_blanco(pdf)

pdf.output(PDF_FILE)
print(f"PDF generado: {PDF_FILE} - {pdf.page_no()} páginas, "
      f"{PAGE_WIDTH:.2f} x {PAGE_HEIGHT:.2f} mm "
      f"({PAGE_WIDTH / 25.4:.3f}\" x {PAGE_HEIGHT / 25.4:.3f}\") "
      f"{'CON sangrado' if CON_SANGRADO else 'SIN sangrado'}")
print(f"Índice general: {N_PAGINAS_INDICE} páginas (español + inglés). "
      f"Catálogo desde la {min(prov_pages_final.values())}, "
      f"hoteles A-Z en la {pag_hoteles}, poblaciones A-Z en la {pag_poblaciones}.")
