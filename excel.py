import math
import re
import unicodedata

import pandas as pd
from fpdf import FPDF

EXCEL_FILE = "excel1.xlsx"
PDF_FILE = "catalogo_hoteles.pdf"

# Controla si se incluye la portada (portada.jpg). Poner False para saltarla.
SHOW_PORTADA = False


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
    "ALAVA": "Vitoria",
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

# Ordenar por: provincia → ES_CAPITAL (True primero) → localidad → estrellas descendentes
df = df.sort_values(
    by=["PROVINCIA", "ES_CAPITAL", "LOCALIDAD", "ESTRELLAS"],
    key=lambda col: (
        col.map(normalizar_provincia) if col.name in ["PROVINCIA", "LOCALIDAD"] else col
    ),
    ascending=[True, False, True, False],
)


# --- PDF ---
class PDF(FPDF):
    def header(self):
        # Encabezado por provincia
        if getattr(self, "provincia_actual", "") not in [None, "", False]:
            self.set_font("Helvetica", "B", 14)
            self.set_text_color(40, 40, 160)
            # Agregar "(cont)" si esta es una página de continuación
            provincia_text = f"PROVINCIA DE {self.provincia_actual.upper()}"
            if getattr(self, "provincia_continuacion", False):
                provincia_text += " (cont.)"
            self.cell(
                0,
                8,
                provincia_text,
                ln=True,
                align="C",
            )
            self.ln(3)

        # Línea superior decorativa
        if getattr(self, "provincia_actual", "") not in [None, "", False]:
            self.set_draw_color(180, 180, 180)
            self.line(10, 20, 200, 20)
            self.ln(5)

    def footer(self):
        if getattr(self, "provincia_actual", "") is None:
            return
        self.set_y(-15)
        self.set_font("Helvetica", "I", 9)
        self.set_text_color(128)
        self.cell(0, 10, f"{self.page_no()}", align="C")


# --- Función para calcular altura real de UNA LÍNEA ---
def calcular_altura_linea(pdf, texto, ancho_efectivo, alto_linea):
    """Calcula cuántas líneas ocupa un texto dado el ancho disponible."""
    if not texto:
        return 0
    w = pdf.get_string_width(texto)
    num_lineas = max(1, math.ceil(w / ancho_efectivo))
    return num_lineas * alto_linea


# --- Función para calcular altura real del bloque completo ---
def calcular_altura_bloque(pdf, lineas_list, ancho_efectivo, alto_linea):
    """Calcula la altura total de un bloque con varias líneas."""
    total_altura = 2  # pequeño margen al inicio
    for linea in lineas_list:
        if linea:
            w = pdf.get_string_width(linea)
            num_lineas = max(1, math.ceil(w / ancho_efectivo))
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


# CONFIGURACIÓN DE GRID FLEXIBLE
COLS = 3
PAGE_WIDTH = 210
MARGIN = 10
COLUMN_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / COLS
Y_START = 30
Y_LIMIT = 280

x_positions = [MARGIN + i * COLUMN_WIDTH for i in range(COLS)]

line_height = 4
ancho_texto = COLUMN_WIDTH - 8

# Obtener lista única de provincias en orden alfabético (sin tildes)
provincias_unicas = sorted(df["PROVINCIA"].unique().tolist(), key=normalizar_provincia)

# Estructura para guardar índice de provincias y sus páginas
indice_provincias = []

# Rellenar indice_provincias con capitales del diccionario
for prov in provincias_unicas:
    prov_normalizada = normalizar_provincia(prov).replace(" ", "")
    capital = CAPITALES.get(prov_normalizada, "-")
    indice_provincias.append({"provincia": prov, "capital": capital, "pagina": None})

# Las páginas reales de cada provincia se registrarán durante la generación
# del catálogo (ver más abajo). El diccionario provincia_pagina_real
# almacenará provincia → número de página real del PDF.
provincia_pagina_real = {}

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

# Añadir segunda página con imagen
try:
    pdf.add_page()
    PAGE_W = pdf.w
    PAGE_H = pdf.h
    pdf.image("Segunda-pagina.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
except Exception as e:
    print(f"No se pudo cargar Segunda-pagina.jpg: {e}")

# --- MARCADOR: página donde empezará el índice de provincias (se insertará al final) ---
# La portada e índice de provincias se generan DESPUÉS del catálogo para poder
# usar las páginas reales. Al final se reordenan las páginas del PDF.

# --- GENERAR CATÁLOGO (reset de variables) ---
x_positions = [MARGIN + i * COLUMN_WIDTH for i in range(COLS)]
pdf.provincia_actual = ""
y_actual = [Y_START] * COLS

provincia_anterior = ""
localidad_anterior = ""
current_col = 0

# Diccionario para rastrear página de cada hotel (sin duplicados)
hotel_pages = {}

for idx, row in df.iterrows():

    provincia = str(row["PROVINCIA"])
    localidad = str(row["LOCALIDAD"])
    hotel_name = str(row["NOMBRE DE EMPRESA"]).strip()

    # CAMBIO DE PROVINCIA → NUEVA PÁGINA Y RESET DE ALTURAS
    if provincia != provincia_anterior:
        provincia_anterior = provincia
        localidad_anterior = ""
        pdf.provincia_actual = provincia
        pdf.provincia_continuacion = False  # Primera página de la provincia
        pdf.add_page()
        current_col = 0
        y_actual = [Y_START] * COLS
        # Registrar página real de esta provincia
        if provincia not in provincia_pagina_real:
            provincia_pagina_real[provincia] = pdf.page_no()

    # ---- REGISTRAR HOTEL CON SU PÁGINA ACTUAL ----
    hotel_name_display = limpiar_nombre_hotel(hotel_name)

    if hotel_name_display and hotel_name_display not in hotel_pages:
        hotel_pages[hotel_name_display] = pdf.page_no()

    # CONSTRUIR TEXTO CON FORMATO (negrita para nombre del hotel)
    # Línea 1: Clasificación + habitaciones (solo si tienen valor real)
    _clasif = str(row["CLASIFICACION HOTEL"]).strip()
    _hab = str(row["NRO. HABITACIONES"]).strip()

    _clasif_ok = _clasif not in ("", "-", "nan", "NaN", "?")
    _hab_ok = _hab not in ("", "-", "nan", "NaN", "?") and _hab.replace(".", "").isdigit()

    _partes = []
    if _clasif_ok:
        _partes.append(_clasif)
    if _hab_ok:
        _partes.append(f"{_hab} Habitaciones")
    linea1 = " · ".join(_partes)
    # Línea 2: Nombre del hotel
    linea2 = limpiar_nombre_hotel(row["NOMBRE DE EMPRESA"])
    linea3 = corregir_preposiciones(row["DIRECCION"])
    linea4 = corregir_preposiciones(f"{row['CP']} {row['LOCALIDAD']}")
    _tel = str(row["TELEFONO1"]).strip()
    _tel_ok = _tel not in ("", "-", "nan", "NaN", "?")
    linea5 = f"Tel: {_tel}".title() if _tel_ok else ""
    linea6 = f"{row['SITIO WEB']}".lower()

    # Convertir a latin-1
    linea1 = linea1.encode("latin-1", "ignore").decode("latin-1")
    linea2 = linea2.encode("latin-1", "ignore").decode("latin-1")
    linea3 = linea3.encode("latin-1", "ignore").decode("latin-1")
    linea4 = linea4.encode("latin-1", "ignore").decode("latin-1")
    linea5 = linea5.encode("latin-1", "ignore").decode("latin-1")
    linea6 = linea6.encode("latin-1", "ignore").decode("latin-1")

    lineas_hotel = [linea2, linea1, linea3, linea4, linea5, linea6]

    # Calcular altura total del hotel
    altura_hotel = calcular_altura_bloque(pdf, lineas_hotel, ancho_texto, line_height)

    # Detectar si la localidad cambió
    hay_cambio_localidad = localidad != localidad_anterior

    # Calcular altura de la localidad si es nueva
    altura_localidad = 0
    if hay_cambio_localidad:
        altura_localidad = (
            calcular_altura_linea(pdf, localidad.upper(), COLUMN_WIDTH - 4, line_height)
            + 4
        )

    # VERIFICAR SI CABE EN LA COLUMNA ACTUAL (localidad + hotel)
    altura_total_requerida = altura_localidad + altura_hotel + 2

    if y_actual[current_col] + altura_total_requerida > Y_LIMIT:
        # NO CABE → pasar a la siguiente columna
        current_col += 1

        # Si se pasa el número de columnas, nueva página
        if current_col >= COLS:
            pdf.provincia_continuacion = True  # Página de continuación de la provincia
            pdf.add_page()
            current_col = 0
            y_actual = [Y_START] * COLS
            # NO resetear localidad_anterior: la localidad solo cambia con provincia

    x = x_positions[current_col]
    y_pos = y_actual[current_col]

    # IMPRIMIR TÍTULO DE LOCALIDAD (si es nueva)
    if hay_cambio_localidad:
        # Aumentar espaciado ANTES de imprimir la nueva localidad (separación del hotel anterior)
        y_pos = y_pos + 1
        localidad_anterior = localidad
        pdf.set_xy(x + 2, y_pos)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(40, 40, 160)
        pdf.multi_cell(
            COLUMN_WIDTH - 4, line_height, localidad.upper(), border=0, align="L"
        )
        # Actualizar y_pos con la nueva altura después de imprimir la localidad
        y_pos = pdf.get_y()
        y_actual[current_col] = y_pos

    # IMPRIMIR TEXTO DEL HOTEL CON FORMATO SEPARADO
    pdf.set_xy(x + 2, y_pos)
    pdf.set_text_color(0, 0, 0)

    # Línea 1: NEGRITA (nombre del hotel)
    pdf.set_font("Helvetica", "B", 9)
    pdf.multi_cell(ancho_texto, line_height, linea2, border=0, align="L")

    # Línea 2: NORMAL (clasificación y habitaciones) — solo si hay dato
    if linea1:
        pdf.set_x(x + 2)
        pdf.set_font("Helvetica", "", 8)
        pdf.multi_cell(ancho_texto, line_height, linea1, border=0, align="L")

    # Líneas 3-6: NORMAL (resetear font siempre, por si linea1 fue omitida)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_x(x + 2)
    pdf.multi_cell(ancho_texto, line_height, linea3, border=0, align="L")
    pdf.set_x(x + 2)
    pdf.multi_cell(ancho_texto, line_height, linea4, border=0, align="L")
    if linea5:
        pdf.set_x(x + 2)
        pdf.multi_cell(ancho_texto, line_height, linea5, border=0, align="L")
    pdf.set_x(x + 2)
    pdf.multi_cell(ancho_texto, line_height, linea6, border=0, align="L")

    # Actualizar y_actual[current_col] con la nueva posición Y después del hotel
    y_actual[current_col] = pdf.get_y() + 2  # pequeño espaciado entre hoteles

# --- GENERAR PORTADA E ÍNDICE DE PROVINCIAS CON PÁGINAS REALES ---
# Actualizar indice_provincias con las páginas reales capturadas durante el catálogo
for item in indice_provincias:
    prov = item["provincia"]
    if prov in provincia_pagina_real:
        item["pagina"] = provincia_pagina_real[prov]

# Anotar en qué página del PDF empieza el bloque de índice de provincias
pagina_inicio_indice_provincias = pdf.page_no() + 1

# --- PORTADA ÍNDICE DE PROVINCIAS ---
pdf.provincia_actual = None
pdf.add_page()

# Fondo azul completo
pdf.set_fill_color(40, 40, 160)
pdf.rect(0, 0, PAGE_WIDTH, 297, "F")

# Línea decorativa superior (blanca, fina)
pdf.set_draw_color(255, 255, 255)
pdf.set_line_width(0.4)
pdf.line(20, 55, PAGE_WIDTH - 20, 55)

# Línea decorativa inferior (blanca, fina)
pdf.line(20, 242, PAGE_WIDTH - 20, 242)

# Rectángulo decorativo central (borde blanco)
pdf.set_line_width(1.2)
pdf.rect(18, 53, PAGE_WIDTH - 36, 191, "D")

# Título principal (blanco)
pdf.set_text_color(255, 255, 255)
pdf.set_font("Helvetica", "B", 28)
pdf.set_xy(0, 90)
pdf.cell(PAGE_WIDTH, 14, "ÍNDICE DE PROVINCIAS", align="C", ln=True)
pdf.set_font("Helvetica", "B", 28)
pdf.set_x(0)
pdf.cell(PAGE_WIDTH, 14, "Y SUS CAPITALES", align="C", ln=True)

# Separador central pequeño
pdf.ln(6)
pdf.set_draw_color(255, 255, 255)
pdf.set_line_width(0.5)
pdf.line(PAGE_WIDTH / 2 - 30, pdf.get_y(), PAGE_WIDTH / 2 + 30, pdf.get_y())
pdf.ln(8)

# Subtítulo en inglés
pdf.set_font("Helvetica", "I", 14)
pdf.set_x(0)
pdf.cell(PAGE_WIDTH, 8, "Index of Provinces and Their Capitals", align="C", ln=True)
pdf.ln(6)

# Descripción
pdf.set_font("Helvetica", "", 11)
pdf.set_x(30)
pdf.multi_cell(
    PAGE_WIDTH - 60,
    7,
    "Provincias de España y sus capitales con página de referencia",
    align="C",
)
pdf.set_x(30)
pdf.multi_cell(
    PAGE_WIDTH - 60,
    7,
    "Provinces of Spain and their capitals with reference page",
    align="C",
)

# País en la parte inferior
pdf.set_font("Helvetica", "B", 13)
pdf.set_xy(0, 252)
pdf.cell(PAGE_WIDTH, 8, "ESPAÑA", align="C", ln=True)

# Resetear colores
pdf.set_text_color(0, 0, 0)
pdf.set_draw_color(0, 0, 0)
pdf.set_line_width(0.2)

# --- PÁGINA DE ÍNDICE DE PROVINCIAS CON PÁGINAS REALES ---
pdf.provincia_actual = None
pdf.add_page()
pdf.set_font("Helvetica", "B", 16)
pdf.set_text_color(0, 0, 0)
pdf.cell(0, 12, "PROVINCIAS DE ESPAÑA Y SUS CAPITALES", ln=True, align="C")
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 10, "PROVINCES OF SPAIN AND THEIR CAPITALS", ln=True, align="C")
pdf.ln(8)

# Tabla de índice en dos columnas paralelas: PROVINCIAS | CAPITALES | Pág.
usable_width_prov = PAGE_WIDTH - 2 * MARGIN
separation_prov = 10
table_width_prov = (usable_width_prov - separation_prov) / 2
col_widths_prov = [table_width_prov * 0.40, table_width_prov * 0.45, table_width_prov * 0.15]
x_left_prov = MARGIN
x_right_prov = MARGIN + table_width_prov + separation_prov
row_h_prov = 8

# Preparar listas izquierda/derecha
n_prov = len(indice_provincias)
mid_prov = (n_prov + 1) // 2
left_items_prov = indice_provincias[:mid_prov]
right_items_prov = indice_provincias[mid_prov:]

while len(left_items_prov) < len(right_items_prov):
    left_items_prov.append({"provincia": "", "capital": "", "pagina": None})
while len(right_items_prov) < len(left_items_prov):
    right_items_prov.append({"provincia": "", "capital": "", "pagina": None})

# Encabezados
pdf.set_font("Helvetica", "B", 10)
y_header_prov = pdf.get_y()

pdf.set_xy(x_left_prov, y_header_prov)
pdf.cell(col_widths_prov[0], row_h_prov, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths_prov[1], row_h_prov, "CAPITALES", border=1, align="C")
pdf.cell(col_widths_prov[2], row_h_prov, "Pág.", border=1, align="C")

pdf.set_xy(x_right_prov, y_header_prov)
pdf.cell(col_widths_prov[0], row_h_prov, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths_prov[1], row_h_prov, "CAPITALES", border=1, align="C")
pdf.cell(col_widths_prov[2], row_h_prov, "Pág.", border=1, align="C")

pdf.set_y(y_header_prov + row_h_prov)

# Filas
pdf.set_font("Helvetica", "", 10)
for i in range(len(left_items_prov)):
    left_p = left_items_prov[i]
    right_p = (
        right_items_prov[i]
        if i < len(right_items_prov)
        else {"provincia": "", "capital": "", "pagina": None}
    )

    prov_l = left_p["provincia"]
    prov_r = right_p["provincia"]

    if not prov_l and not prov_r:
        continue

    y_p = pdf.get_y()

    capital_l = left_p["capital"]
    page_l = str(left_p["pagina"]) if left_p["pagina"] is not None else "..."
    pdf.set_xy(x_left_prov, y_p)
    pdf.cell(
        col_widths_prov[0],
        row_h_prov,
        prov_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths_prov[1],
        row_h_prov,
        capital_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths_prov[2], row_h_prov, page_l, border=1, align="C")

    capital_r = right_p["capital"]
    page_r = str(right_p["pagina"]) if right_p["pagina"] is not None else "..."
    pdf.set_xy(x_right_prov, y_p)
    pdf.cell(
        col_widths_prov[0],
        row_h_prov,
        prov_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths_prov[1],
        row_h_prov,
        capital_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths_prov[2], row_h_prov, page_r, border=1, align="C")

    pdf.set_y(y_p + row_h_prov)

# Anotar la última página del bloque de índice de provincias
pagina_fin_indice_provincias = pdf.page_no()

# --- REORDENAR PÁGINAS: mover el bloque del índice de provincias a su posición original ---
# pdf.pages es un dict {1: contenido, 2: contenido, ...}
# Reconstruimos el dict en el nuevo orden e reindexamos desde 1.
paginas_previas = 2 if SHOW_PORTADA else 1  # segunda-pagina (+ portada opcional)
posicion_destino = paginas_previas + 1       # primera página donde irá el índice de provincias

total_paginas = pdf.page_no()

# Construir lista ordenada de contenidos en el orden original (índices 1..N)
orden_original = [pdf.pages[p] for p in range(1, total_paginas + 1)]

# Extraer el bloque del índice de provincias
bloque_indice = orden_original[pagina_inicio_indice_provincias - 1 : pagina_fin_indice_provincias]

# Lista sin ese bloque
resto = (
    orden_original[: pagina_inicio_indice_provincias - 1]
    + orden_original[pagina_fin_indice_provincias :]
)

# Insertar el bloque en la posición destino (base-0 en la lista)
nueva_orden = (
    resto[: posicion_destino - 1]
    + bloque_indice
    + resto[posicion_destino - 1 :]
)

# Reconstruir pdf.pages con el nuevo orden
pdf.pages = {i + 1: contenido for i, contenido in enumerate(nueva_orden)}

# --- PORTADA ÍNDICE ALFABÉTICO DE HOTELES ---
pdf.provincia_actual = None
pdf.add_page()

# Fondo azul completo
pdf.set_fill_color(40, 40, 160)
pdf.rect(0, 0, PAGE_WIDTH, 297, "F")

# Línea decorativa superior (blanca, fina)
pdf.set_draw_color(255, 255, 255)
pdf.set_line_width(0.4)
pdf.line(20, 55, PAGE_WIDTH - 20, 55)

# Línea decorativa inferior (blanca, fina)
pdf.line(20, 242, PAGE_WIDTH - 20, 242)

# Rectángulo decorativo central (borde blanco)
pdf.set_line_width(1.2)
pdf.rect(18, 53, PAGE_WIDTH - 36, 191, "D")

# Título principal (blanco)
pdf.set_text_color(255, 255, 255)
pdf.set_font("Helvetica", "B", 28)
pdf.set_xy(0, 90)
pdf.cell(PAGE_WIDTH, 14, "ÍNDICE ALFABÉTICO", align="C", ln=True)
pdf.set_font("Helvetica", "B", 28)
pdf.set_x(0)
pdf.cell(PAGE_WIDTH, 14, "DE HOTELES", align="C", ln=True)

# Separador central pequeño
pdf.ln(6)
pdf.set_draw_color(255, 255, 255)
pdf.set_line_width(0.5)
pdf.line(PAGE_WIDTH / 2 - 30, pdf.get_y(), PAGE_WIDTH / 2 + 30, pdf.get_y())
pdf.ln(8)

# Subtítulo en inglés
pdf.set_font("Helvetica", "I", 14)
pdf.set_x(0)
pdf.cell(PAGE_WIDTH, 8, "Alphabetical Index of Hotels", align="C", ln=True)
pdf.ln(6)

# Descripción
pdf.set_font("Helvetica", "", 11)
pdf.set_x(30)
pdf.multi_cell(
    PAGE_WIDTH - 60,
    7,
    "Hoteles legalmente autorizados existentes en España",
    align="C",
)
pdf.set_x(30)
pdf.multi_cell(
    PAGE_WIDTH - 60,
    7,
    "Hotels legally authorized existing in Spain",
    align="C",
)

# Año en la parte inferior
pdf.set_font("Helvetica", "B", 13)
pdf.set_xy(0, 252)
pdf.cell(PAGE_WIDTH, 8, "ESPAÑA", align="C", ln=True)

# Resetear colores
pdf.set_text_color(0, 0, 0)
pdf.set_draw_color(0, 0, 0)
pdf.set_line_width(0.2)

# --- INICIAR ÍNDICE ALFABÉTICO DE HOTELES ---
pdf.provincia_actual = None
pdf.add_page()

# Títulos (sin línea separadora)
pdf.set_font("Helvetica", "B", 12)
pdf.set_text_color(0, 0, 0)
pdf.cell(
    0,
    8,
    "Hoteles legalmente autorizados existentes en España, por orden alfabético.",
    ln=True,
    align="C",
)
pdf.cell(
    0,
    8,
    "Hotels legally authorized existing in Spain, in alphabetical order.",
    ln=True,
    align="C",
)
pdf.ln(4)

# Lista de hoteles ordenada
hoteles_lista = sorted(hotel_pages.keys(), key=lambda x: x.lower())

# Configuración: 3 columnas verticales
COLS_INDEX = 3
col_width_index = (PAGE_WIDTH - 2 * MARGIN) / COLS_INDEX
row_height_index = 4.5  # Ajustado para tipografía pequeña
y_start_index = pdf.get_y()
y_limit_index = 280


# ---- FUNCIÓN DE FORMATO (tipografía 6pt equivalente) ----
def format_index_entry(pdf, name, page, max_width):
    encoded_name = name.encode("latin-1", "ignore").decode("latin-1")
    page_str = str(page)

    # Reservar espacio para número de página
    space_reserved = pdf.get_string_width(page_str) + 1.5
    max_name_width = max_width - space_reserved - 2

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
pdf.set_font("Helvetica", "", 6.5)
pdf.set_text_color(0, 0, 0)

pdf.set_y(y_start_index)

x_cols = [MARGIN + i * col_width_index for i in range(COLS_INDEX)]
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
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(
                0,
                8,
                "Hoteles legalmente autorizados existentes en España, por orden alfabético.",
                ln=True,
                align="C",
            )
            pdf.cell(
                0,
                8,
                "Hotels legally authorized existing in Spain, in alphabetical order.",
                ln=True,
                align="C",
            )
            pdf.ln(4)
            pdf.set_font("Helvetica", "", 6.5)

            current_col = 0
            y_cols = [pdf.get_y()] * COLS_INDEX
            page_count += 1

    # Imprimir hotel en columna actual
    hotel = hoteles_lista[hotel_idx]
    pagina = hotel_pages[hotel]
    linea = format_index_entry(pdf, hotel, pagina, col_width_index - 5)

    pdf.set_xy(x_cols[current_col] + 1.5, y_cols[current_col])
    pdf.cell(col_width_index - 3, row_height_index, linea, border=0, align="L")

    y_cols[current_col] += row_height_index
    hotel_idx += 1

# --- PORTADA ÍNDICE ALFABÉTICO DE POBLACIONES ---
pdf.provincia_actual = None
pdf.add_page()

# Fondo azul completo
pdf.set_fill_color(40, 40, 160)
pdf.rect(0, 0, PAGE_WIDTH, 297, "F")

# Línea decorativa superior (blanca, fina)
pdf.set_draw_color(255, 255, 255)
pdf.set_line_width(0.4)
pdf.line(20, 55, PAGE_WIDTH - 20, 55)

# Línea decorativa inferior (blanca, fina)
pdf.line(20, 242, PAGE_WIDTH - 20, 242)

# Rectángulo decorativo central (borde blanco)
pdf.set_line_width(1.2)
pdf.rect(18, 53, PAGE_WIDTH - 36, 191, "D")

# Título principal (blanco)
pdf.set_text_color(255, 255, 255)
pdf.set_font("Helvetica", "B", 28)
pdf.set_xy(0, 90)
pdf.cell(PAGE_WIDTH, 14, "ÍNDICE ALFABÉTICO", align="C", ln=True)
pdf.set_font("Helvetica", "B", 28)
pdf.set_x(0)
pdf.cell(PAGE_WIDTH, 14, "DE POBLACIONES", align="C", ln=True)

# Separador central pequeño
pdf.ln(6)
pdf.set_draw_color(255, 255, 255)
pdf.set_line_width(0.5)
pdf.line(PAGE_WIDTH / 2 - 30, pdf.get_y(), PAGE_WIDTH / 2 + 30, pdf.get_y())
pdf.ln(8)

# Subtítulo en inglés
pdf.set_font("Helvetica", "I", 14)
pdf.set_x(0)
pdf.cell(PAGE_WIDTH, 8, "Alphabetical Index of Towns", align="C", ln=True)
pdf.ln(6)

# Descripción
pdf.set_font("Helvetica", "", 11)
pdf.set_x(30)
pdf.multi_cell(
    PAGE_WIDTH - 60,
    7,
    "Poblaciones de España con hoteles legalmente autorizados",
    align="C",
)
pdf.set_x(30)
pdf.multi_cell(
    PAGE_WIDTH - 60,
    7,
    "Spanish towns with legally authorized hotels",
    align="C",
)

# País en la parte inferior
pdf.set_font("Helvetica", "B", 13)
pdf.set_xy(0, 252)
pdf.cell(PAGE_WIDTH, 8, "ESPAÑA", align="C", ln=True)

# Resetear colores
pdf.set_text_color(0, 0, 0)
pdf.set_draw_color(0, 0, 0)
pdf.set_line_width(0.2)

# --- INICIAR ÍNDICE ALFABÉTICO DE POBLACIONES ---
pdf.provincia_actual = None
pdf.add_page()

# Títulos del índice de poblaciones
pdf.set_font("Helvetica", "B", 12)
pdf.set_text_color(0, 0, 0)
pdf.cell(
    0,
    8,
    "Poblaciones de España con hoteles legalmente autorizados, por orden alfabético.",
    ln=True,
    align="C",
)
pdf.cell(
    0,
    8,
    "Spanish towns with legally authorized hotels, in alphabetical order.",
    ln=True,
    align="C",
)
pdf.ln(4)

# Construir diccionario de poblaciones → página (primera aparición)
poblacion_pages = {}
for idx, row in df.iterrows():
    localidad = str(row["LOCALIDAD"]).strip()
    hotel_name_r = str(row["NOMBRE DE EMPRESA"]).strip()
    hotel_name_display_r = limpiar_nombre_hotel(hotel_name_r)
    if localidad and localidad not in poblacion_pages:
        # Obtener la página del primer hotel de esa localidad
        if hotel_name_display_r in hotel_pages:
            poblacion_pages[localidad] = hotel_pages[hotel_name_display_r]

# Lista de poblaciones ordenada alfabéticamente (sin tildes)
poblaciones_lista = sorted(poblacion_pages.keys(), key=lambda x: normalizar_ciudad(x))

# Configuración: 3 columnas verticales (igual que el índice de hoteles)
COLS_POB = 3
col_width_pob = (PAGE_WIDTH - 2 * MARGIN) / COLS_POB
row_height_pob = 4.5
y_start_pob = pdf.get_y()
y_limit_pob = 280

# ---- IMPRIMIR ÍNDICE DE POBLACIONES EN COLUMNAS VERTICALES ----
pdf.set_font("Helvetica", "", 6.5)
pdf.set_text_color(0, 0, 0)
pdf.set_y(y_start_pob)

x_cols_pob = [MARGIN + i * col_width_pob for i in range(COLS_POB)]
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
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(
                0,
                8,
                "Poblaciones de España con hoteles legalmente autorizados, por orden alfabético.",
                ln=True,
                align="C",
            )
            pdf.cell(
                0,
                8,
                "Spanish towns with legally authorized hotels, in alphabetical order.",
                ln=True,
                align="C",
            )
            pdf.ln(4)
            pdf.set_font("Helvetica", "", 6.5)

            current_col_pob = 0
            y_cols_pob = [pdf.get_y()] * COLS_POB

    poblacion = poblaciones_lista[pob_idx]
    pagina_pob = poblacion_pages[poblacion]
    linea_pob = format_index_entry(pdf, poblacion, pagina_pob, col_width_pob - 5)

    pdf.set_xy(x_cols_pob[current_col_pob] + 1.5, y_cols_pob[current_col_pob])
    pdf.cell(col_width_pob - 3, row_height_pob, linea_pob, border=0, align="L")

    y_cols_pob[current_col_pob] += row_height_pob
    pob_idx += 1

pdf.output(PDF_FILE)
print("PDF generado con índice alfabético de 5 columnas verticales:", PDF_FILE)