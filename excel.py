import math
import pandas as pd
from fpdf import FPDF

EXCEL_FILE = "excel1.xlsx"
PDF_FILE = "catalogo_hoteles.pdf"

# Controla si se incluye la portada (portada.jpg). Poner False para saltarla.
SHOW_PORTADA = False

import unicodedata


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


# Extraer valor numérico de la clasificación para ordenar por estrellas (5->0)
def extraer_estrellas(val):
    try:
        s = str(val)
        # Buscar un número al inicio o en cualquier parte
        import re

        m = re.search(r"(\d+)", s)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return 0


df["ESTRELLAS"] = df["CLASIFICACION HOTEL"].apply(extraer_estrellas)


# Extraer estrellas ya asignadas
# Crear función de normalización robusta para localidades/capitales
def normalizar_ciudad(nombre):
    if not isinstance(nombre, str):
        nombre = str(nombre)
    # Eliminar tildes y normalizar unicode
    nfkd = unicodedata.normalize("NFKD", nombre)
    sin_tildes = "".join(c for c in nfkd if not unicodedata.combining(c))
    # Sustituir separadores comunes por espacios
    s = sin_tildes.replace("/", " ").replace("-", " ")
    # Colapsar espacios duplicados y recortar
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
        import re

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


# Crear PDF
pdf = PDF(format=(152.4, 228.6))
pdf.set_auto_page_break(auto=False)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(0, 0, 0)
pdf.provincia_continuacion = False

# --- GENERAR PÁGINA DE ÍNDICE ---
# Obtener lista única de provincias en orden alfabético (sin tildes)
provincias_unicas = sorted(df["PROVINCIA"].unique().tolist(), key=normalizar_provincia)

# Estructura para guardar índice de provincias y sus páginas
indice_provincias = []

# Crear página de índice
pdf.add_page()
pdf.set_font("Helvetica", "B", 16)
pdf.set_text_color(0, 0, 0)
pdf.cell(0, 12, "PROVINCIAS DE ESPAÑA Y SUS CAPITALES", ln=True, align="C")
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 10, "PROVINCES OF SPAIN AND THEIR CAPITALS", ln=True, align="C")
pdf.ln(8)

# Tabla de índice en dos columnas paralelas: PROVINCIAS | CAPITALES | Pág.
page_width_local = 210
margin_local = 10
usable_width = page_width_local - 2 * margin_local
separation = 10
table_width = (usable_width - separation) / 2
col_widths = [table_width * 0.40, table_width * 0.45, table_width * 0.15]
x_left = margin_local
x_right = margin_local + table_width + separation
row_h = 8

# Preparar listas izquierda/derecha
n = len(provincias_unicas)
mid = (n + 1) // 2
left_provs = provincias_unicas[:mid]
right_provs = provincias_unicas[mid:]

# Rellenar indice_provincias con capitales del diccionario
for prov in provincias_unicas:
    prov_normalizada = normalizar_provincia(prov).replace(" ", "")
    capital = CAPITALES.get(prov_normalizada, "-")
    indice_provincias.append({"provincia": prov, "capital": capital, "pagina": None})

# Encabezados para ambas tablas (alineados a la misma altura)
pdf.set_font("Helvetica", "B", 10)
y_header = pdf.get_y()

# Imprimir encabezado izquierdo (LEVANTADO)
pdf.set_xy(x_left, y_header - row_h)  # Restar row_h para levantarlo
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Imprimir encabezado derecho (MISMA ALTURA levantada)
pdf.set_xy(x_right, y_header - row_h)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# La posición actual se mantiene en y_header para que las filas comiencen aquí
# NO avanzamos, las filas empiezan justo donde estaba antes la cabecera

# Imprimir filas en paralelo (sin solapamiento)
pdf.set_font("Helvetica", "", 10)
for i in range(len(left_provs)):
    prov_l = left_provs[i]
    capital_l = CAPITALES.get(normalizar_provincia(prov_l).replace(" ", ""), "-")

    prov_r = right_provs[i] if i < len(right_provs) else ""
    capital_r = (
        CAPITALES.get(normalizar_provincia(prov_r).replace(" ", ""), "-")
        if prov_r
        else ""
    )

    y = pdf.get_y()

    # Imprimir fila izquierda
    pdf.set_xy(x_left, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, "...", border=1, align="C")

    # Imprimir fila derecha (MISMO Y que la izquierda)
    pdf.set_xy(x_right, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, "...", border=1, align="C")

    # Avanzar Y solo una vez (ambas columnas se imprimen al mismo nivel)
    pdf.set_y(y + row_h)

# CONFIGURACIÓN DE GRID FLEXIBLE
COLS = 3
PAGE_WIDTH = 210
MARGIN = 10
COLUMN_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / COLS
Y_START = 30
Y_LIMIT = 280

x_positions = [MARGIN + i * COLUMN_WIDTH for i in range(COLS)]
pdf.provincia_actual = ""

line_height = 4
ancho_texto = COLUMN_WIDTH - 8

# Array para rastrear la altura Y actual en cada columna
y_actual = [Y_START] * COLS

# GENERAR FICHAS
provincia_anterior = ""
localidad_anterior = ""
current_col = 0
indice_actual = 0

# Calcular números de página para cada provincia
# Simular el bucle de generación para saber cuándo se añade página
pagina_actual = 3  # Primera página de catálogo (1 portada, 2 índice, 3 catálogo)
columna_actual = 0
altura_y = 30
provincia_anterior_temp = ""

for idx, row in df.iterrows():
    provincia = str(row["PROVINCIA"])

    if provincia != provincia_anterior_temp:
        # Nueva provincia: encontrar en indice_provincias y asignar página
        for item in indice_provincias:
            if item["provincia"] == provincia and item["pagina"] is None:
                item["pagina"] = pagina_actual
                break

        # Actualizar valores para nueva provincia
        provincia_anterior_temp = provincia
        pagina_actual += 1
        columna_actual = 0
        altura_y = 30

# Ahora que tenemos los números de página, regenerar el PDF con índice correcto
# Crear un nuevo PDF desde cero
pdf = PDF()
pdf.set_auto_page_break(auto=False)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(0, 0, 0)

# Añadir portada a toda la página si existe
if SHOW_PORTADA:
    try:
        pdf.add_page()
        PAGE_W = pdf.w  # ancho de la página en mm
        PAGE_H = pdf.h  # alto de la página en mm
        pdf.image("portada.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
    except Exception as e:
        # Si falla la carga, seguimos sin portada
        print(f"No se pudo cargar portada.jpg: {e}")

# Añadir segunda página con imagen
try:
    pdf.add_page()
    PAGE_W = pdf.w
    PAGE_H = pdf.h
    pdf.image("Segunda-pagina.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
except Exception as e:
    print(f"No se pudo cargar Segunda-pagina.jpg: {e}")

# --- REGENERAR PÁGINA DE ÍNDICE CON NÚMEROS CORRECTOS ---
pdf.add_page()
pdf.set_font("Helvetica", "B", 16)
pdf.set_text_color(0, 0, 0)
pdf.cell(0, 12, "PROVINCIAS DE ESPAÑA Y SUS CAPITALES", ln=True, align="C")
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 10, "PROVINCES OF SPAIN AND THEIR CAPITALS", ln=True, align="C")
pdf.ln(8)

# Tabla de índice en dos columnas paralelas: PROVINCIAS | CAPITALES | Pág.
usable_width = PAGE_WIDTH - 2 * MARGIN
separation = 10
table_width = (usable_width - separation) / 2
col_widths = [table_width * 0.40, table_width * 0.45, table_width * 0.15]
x_left = MARGIN
x_right = MARGIN + table_width + separation
row_h = 8

# Preparar listas izquierda/derecha
n = len(indice_provincias)
mid = (n + 1) // 2
left_items = indice_provincias[:mid]
right_items = indice_provincias[mid:]

# Igualar longitudes para evitar filas vacías al final
while len(left_items) < len(right_items):
    left_items.append({"provincia": "", "capital": "", "pagina": None})
while len(right_items) < len(left_items):
    right_items.append({"provincia": "", "capital": "", "pagina": None})

# Encabezados para ambas tablas EN POSICIÓN NATURAL
pdf.set_font("Helvetica", "B", 10)
y_header_start = pdf.get_y()

# Imprimir encabezado izquierdo
pdf.set_xy(x_left, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Imprimir encabezado derecho (MISMA ALTURA)
pdf.set_xy(x_right, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Avanzar posición después de la cabecera
y_first_row = y_header_start + row_h
pdf.set_y(y_first_row)

# Imprimir filas en paralelo (sin filas vacías al final)
pdf.set_font("Helvetica", "", 10)
for i in range(len(left_items)):
    left = left_items[i]
    right = (
        right_items[i]
        if i < len(right_items)
        else {"provincia": "", "capital": "", "pagina": None}
    )

    prov_l = left["provincia"]
    prov_r = right["provincia"]

    # Saltar si ambas provincias están vacías
    if not prov_l and not prov_r:
        continue

    y = pdf.get_y()

    # Imprimir fila izquierda
    capital_l = left["capital"]
    page_l = str(left["pagina"]) if left["pagina"] is not None else "..."
    pdf.set_xy(x_left, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_l, border=1, align="C")

    # Imprimir fila derecha (MISMO Y que la izquierda)
    capital_r = right["capital"]
    page_r = str(right["pagina"]) if right["pagina"] is not None else "..."
    pdf.set_xy(x_right, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_r, border=1, align="C")

    # Avanzar Y para la siguiente fila
    pdf.set_y(y + row_h)

# CONFIGURACIÓN DE GRID FLEXIBLE
COLS = 3
PAGE_WIDTH = 210
MARGIN = 10
COLUMN_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / COLS
Y_START = 30
Y_LIMIT = 280

x_positions = [MARGIN + i * COLUMN_WIDTH for i in range(COLS)]
pdf.provincia_actual = ""

line_height = 4
ancho_texto = COLUMN_WIDTH - 8

# Array para rastrear la altura Y actual en cada columna
y_actual = [Y_START] * COLS

# GENERAR FICHAS
provincia_anterior = ""
localidad_anterior = ""
current_col = 0
indice_actual = 0

# Calcular números de página para cada provincia
# Simular el bucle de generación para saber cuándo se añade página
pagina_actual = 3  # Primera página de catálogo (1 portada, 2 índice, 3 catálogo)
columna_actual = 0
altura_y = 30
provincia_anterior_temp = ""

for idx, row in df.iterrows():
    provincia = str(row["PROVINCIA"])

    if provincia != provincia_anterior_temp:
        # Nueva provincia: encontrar en indice_provincias y asignar página
        for item in indice_provincias:
            if item["provincia"] == provincia and item["pagina"] is None:
                item["pagina"] = pagina_actual
                break

        # Actualizar valores para nueva provincia
        provincia_anterior_temp = provincia
        pagina_actual += 1
        columna_actual = 0
        altura_y = 30

# Ahora que tenemos los números de página, regenerar el PDF con índice correcto
# Crear un nuevo PDF desde cero
pdf = PDF()
pdf.set_auto_page_break(auto=False)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(0, 0, 0)

# Añadir portada a toda la página si existe
if SHOW_PORTADA:
    try:
        pdf.add_page()
        PAGE_W = pdf.w  # ancho de la página en mm
        PAGE_H = pdf.h  # alto de la página en mm
        pdf.image("portada.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
    except Exception as e:
        # Si falla la carga, seguimos sin portada
        print(f"No se pudo cargar portada.jpg: {e}")

# Añadir segunda página con imagen
try:
    pdf.add_page()
    PAGE_W = pdf.w
    PAGE_H = pdf.h
    pdf.image("Segunda-pagina.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
except Exception as e:
    print(f"No se pudo cargar Segunda-pagina.jpg: {e}")

# --- REGENERAR PÁGINA DE ÍNDICE CON NÚMEROS CORRECTOS ---
pdf.add_page()
pdf.set_font("Helvetica", "B", 16)
pdf.set_text_color(0, 0, 0)
pdf.cell(0, 12, "PROVINCIAS DE ESPAÑA Y SUS CAPITALES", ln=True, align="C")
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 10, "PROVINCES OF SPAIN AND THEIR CAPITALS", ln=True, align="C")
pdf.ln(8)

# Tabla de índice en dos columnas paralelas: PROVINCIAS | CAPITALES | Pág.
usable_width = PAGE_WIDTH - 2 * MARGIN
separation = 10
table_width = (usable_width - separation) / 2
col_widths = [table_width * 0.40, table_width * 0.45, table_width * 0.15]
x_left = MARGIN
x_right = MARGIN + table_width + separation
row_h = 8

# Preparar listas izquierda/derecha
n = len(indice_provincias)
mid = (n + 1) // 2
left_items = indice_provincias[:mid]
right_items = indice_provincias[mid:]

# Igualar longitudes para evitar filas vacías al final
while len(left_items) < len(right_items):
    left_items.append({"provincia": "", "capital": "", "pagina": None})
while len(right_items) < len(left_items):
    right_items.append({"provincia": "", "capital": "", "pagina": None})

# Encabezados para ambas tablas EN POSICIÓN NATURAL
pdf.set_font("Helvetica", "B", 10)
y_header_start = pdf.get_y()

# Imprimir encabezado izquierdo
pdf.set_xy(x_left, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Imprimir encabezado derecho (MISMA ALTURA)
pdf.set_xy(x_right, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Avanzar posición después de la cabecera
y_first_row = y_header_start + row_h
pdf.set_y(y_first_row)

# Imprimir filas en paralelo (sin filas vacías al final)
pdf.set_font("Helvetica", "", 10)
for i in range(len(left_items)):
    left = left_items[i]
    right = (
        right_items[i]
        if i < len(right_items)
        else {"provincia": "", "capital": "", "pagina": None}
    )

    prov_l = left["provincia"]
    prov_r = right["provincia"]

    # Saltar si ambas provincias están vacías
    if not prov_l and not prov_r:
        continue

    y = pdf.get_y()

    # Imprimir fila izquierda
    capital_l = left["capital"]
    page_l = str(left["pagina"]) if left["pagina"] is not None else "..."
    pdf.set_xy(x_left, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_l, border=1, align="C")

    # Imprimir fila derecha (MISMO Y que la izquierda)
    capital_r = right["capital"]
    page_r = str(right["pagina"]) if right["pagina"] is not None else "..."
    pdf.set_xy(x_right, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_r, border=1, align="C")

    # Avanzar Y para la siguiente fila
    pdf.set_y(y + row_h)

# CONFIGURACIÓN DE GRID FLEXIBLE
COLS = 3
PAGE_WIDTH = 210
MARGIN = 10
COLUMN_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / COLS
Y_START = 30
Y_LIMIT = 280

x_positions = [MARGIN + i * COLUMN_WIDTH for i in range(COLS)]
pdf.provincia_actual = ""

line_height = 4
ancho_texto = COLUMN_WIDTH - 8

# Array para rastrear la altura Y actual en cada columna
y_actual = [Y_START] * COLS

# GENERAR FICHAS
provincia_anterior = ""
localidad_anterior = ""
current_col = 0
indice_actual = 0

# Calcular números de página para cada provincia
# Simular el bucle de generación para saber cuándo se añade página
pagina_actual = 3  # Primera página de catálogo (1 portada, 2 índice, 3 catálogo)
columna_actual = 0
altura_y = 30
provincia_anterior_temp = ""

for idx, row in df.iterrows():
    provincia = str(row["PROVINCIA"])

    if provincia != provincia_anterior_temp:
        # Nueva provincia: encontrar en indice_provincias y asignar página
        for item in indice_provincias:
            if item["provincia"] == provincia and item["pagina"] is None:
                item["pagina"] = pagina_actual
                break

        # Actualizar valores para nueva provincia
        provincia_anterior_temp = provincia
        pagina_actual += 1
        columna_actual = 0
        altura_y = 30

# Ahora que tenemos los números de página, regenerar el PDF con índice correcto
# Crear un nuevo PDF desde cero
pdf = PDF()
pdf.set_auto_page_break(auto=False)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(0, 0, 0)

# Añadir portada a toda la página si existe
if SHOW_PORTADA:
    try:
        pdf.add_page()
        PAGE_W = pdf.w  # ancho de la página en mm
        PAGE_H = pdf.h  # alto de la página en mm
        pdf.image("portada.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
    except Exception as e:
        # Si falla la carga, seguimos sin portada
        print(f"No se pudo cargar portada.jpg: {e}")

# Añadir segunda página con imagen
try:
    pdf.add_page()
    PAGE_W = pdf.w
    PAGE_H = pdf.h
    pdf.image("Segunda-pagina.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
except Exception as e:
    print(f"No se pudo cargar Segunda-pagina.jpg: {e}")

# --- REGENERAR PÁGINA DE ÍNDICE CON NÚMEROS CORRECTOS ---
pdf.add_page()
pdf.set_font("Helvetica", "B", 16)
pdf.set_text_color(0, 0, 0)
pdf.cell(0, 12, "PROVINCIAS DE ESPAÑA Y SUS CAPITALES", ln=True, align="C")
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 10, "PROVINCES OF SPAIN AND THEIR CAPITALS", ln=True, align="C")
pdf.ln(8)

# Tabla de índice en dos columnas paralelas: PROVINCIAS | CAPITALES | Pág.
usable_width = PAGE_WIDTH - 2 * MARGIN
separation = 10
table_width = (usable_width - separation) / 2
col_widths = [table_width * 0.40, table_width * 0.45, table_width * 0.15]
x_left = MARGIN
x_right = MARGIN + table_width + separation
row_h = 8

# Preparar listas izquierda/derecha
n = len(indice_provincias)
mid = (n + 1) // 2
left_items = indice_provincias[:mid]
right_items = indice_provincias[mid:]

# Igualar longitudes para evitar filas vacías al final
while len(left_items) < len(right_items):
    left_items.append({"provincia": "", "capital": "", "pagina": None})
while len(right_items) < len(left_items):
    right_items.append({"provincia": "", "capital": "", "pagina": None})

# Encabezados para ambas tablas EN POSICIÓN NATURAL
pdf.set_font("Helvetica", "B", 10)
y_header_start = pdf.get_y()

# Imprimir encabezado izquierdo
pdf.set_xy(x_left, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Imprimir encabezado derecho (MISMA ALTURA)
pdf.set_xy(x_right, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Avanzar posición después de la cabecera
y_first_row = y_header_start + row_h
pdf.set_y(y_first_row)

# Imprimir filas en paralelo (sin filas vacías al final)
pdf.set_font("Helvetica", "", 10)
for i in range(len(left_items)):
    left = left_items[i]
    right = (
        right_items[i]
        if i < len(right_items)
        else {"provincia": "", "capital": "", "pagina": None}
    )

    prov_l = left["provincia"]
    prov_r = right["provincia"]

    # Saltar si ambas provincias están vacías
    if not prov_l and not prov_r:
        continue

    y = pdf.get_y()

    # Imprimir fila izquierda
    capital_l = left["capital"]
    page_l = str(left["pagina"]) if left["pagina"] is not None else "..."
    pdf.set_xy(x_left, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_l, border=1, align="C")

    # Imprimir fila derecha (MISMO Y que la izquierda)
    capital_r = right["capital"]
    page_r = str(right["pagina"]) if right["pagina"] is not None else "..."
    pdf.set_xy(x_right, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_r, border=1, align="C")

    # Avanzar Y para la siguiente fila
    pdf.set_y(y + row_h)

# CONFIGURACIÓN DE GRID FLEXIBLE
COLS = 3
PAGE_WIDTH = 210
MARGIN = 10
COLUMN_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / COLS
Y_START = 30
Y_LIMIT = 280

x_positions = [MARGIN + i * COLUMN_WIDTH for i in range(COLS)]
pdf.provincia_actual = ""

line_height = 4
ancho_texto = COLUMN_WIDTH - 8

# Array para rastrear la altura Y actual en cada columna
y_actual = [Y_START] * COLS

# GENERAR FICHAS
provincia_anterior = ""
localidad_anterior = ""
current_col = 0
indice_actual = 0

# Calcular números de página para cada provincia
# Simular el bucle de generación para saber cuándo se añade página
pagina_actual = 3  # Primera página de catálogo (1 portada, 2 índice, 3 catálogo)
columna_actual = 0
altura_y = 30
provincia_anterior_temp = ""

for idx, row in df.iterrows():
    provincia = str(row["PROVINCIA"])

    if provincia != provincia_anterior_temp:
        # Nueva provincia: encontrar en indice_provincias y asignar página
        for item in indice_provincias:
            if item["provincia"] == provincia and item["pagina"] is None:
                item["pagina"] = pagina_actual
                break

        # Actualizar valores para nueva provincia
        provincia_anterior_temp = provincia
        pagina_actual += 1
        columna_actual = 0
        altura_y = 30

# Ahora que tenemos los números de página, regenerar el PDF con índice correcto
# Crear un nuevo PDF desde cero
pdf = PDF()
pdf.set_auto_page_break(auto=False)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(0, 0, 0)

# Añadir portada a toda la página si existe
if SHOW_PORTADA:
    try:
        pdf.add_page()
        PAGE_W = pdf.w  # ancho de la página en mm
        PAGE_H = pdf.h  # alto de la página en mm
        pdf.image("portada.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
    except Exception as e:
        # Si falla la carga, seguimos sin portada
        print(f"No se pudo cargar portada.jpg: {e}")

# Añadir segunda página con imagen
try:
    pdf.add_page()
    PAGE_W = pdf.w
    PAGE_H = pdf.h
    pdf.image("Segunda-pagina.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
except Exception as e:
    print(f"No se pudo cargar Segunda-pagina.jpg: {e}")

# --- REGENERAR PÁGINA DE ÍNDICE CON NÚMEROS CORRECTOS ---
pdf.add_page()
pdf.set_font("Helvetica", "B", 16)
pdf.set_text_color(0, 0, 0)
pdf.cell(0, 12, "PROVINCIAS DE ESPAÑA Y SUS CAPITALES", ln=True, align="C")
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 10, "PROVINCES OF SPAIN AND THEIR CAPITALS", ln=True, align="C")
pdf.ln(8)

# Tabla de índice en dos columnas paralelas: PROVINCIAS | CAPITALES | Pág.
usable_width = PAGE_WIDTH - 2 * MARGIN
separation = 10
table_width = (usable_width - separation) / 2
col_widths = [table_width * 0.40, table_width * 0.45, table_width * 0.15]
x_left = MARGIN
x_right = MARGIN + table_width + separation
row_h = 8

# Preparar listas izquierda/derecha
n = len(indice_provincias)
mid = (n + 1) // 2
left_items = indice_provincias[:mid]
right_items = indice_provincias[mid:]

# Igualar longitudes para evitar filas vacías al final
while len(left_items) < len(right_items):
    left_items.append({"provincia": "", "capital": "", "pagina": None})
while len(right_items) < len(left_items):
    right_items.append({"provincia": "", "capital": "", "pagina": None})

# Encabezados para ambas tablas EN POSICIÓN NATURAL
pdf.set_font("Helvetica", "B", 10)
y_header_start = pdf.get_y()

# Imprimir encabezado izquierdo
pdf.set_xy(x_left, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Imprimir encabezado derecho (MISMA ALTURA)
pdf.set_xy(x_right, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Avanzar posición después de la cabecera
y_first_row = y_header_start + row_h
pdf.set_y(y_first_row)

# Imprimir filas en paralelo (sin filas vacías al final)
pdf.set_font("Helvetica", "", 10)
for i in range(len(left_items)):
    left = left_items[i]
    right = (
        right_items[i]
        if i < len(right_items)
        else {"provincia": "", "capital": "", "pagina": None}
    )

    prov_l = left["provincia"]
    prov_r = right["provincia"]

    # Saltar si ambas provincias están vacías
    if not prov_l and not prov_r:
        continue

    y = pdf.get_y()

    # Imprimir fila izquierda
    capital_l = left["capital"]
    page_l = str(left["pagina"]) if left["pagina"] is not None else "..."
    pdf.set_xy(x_left, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_l, border=1, align="C")

    # Imprimir fila derecha (MISMO Y que la izquierda)
    capital_r = right["capital"]
    page_r = str(right["pagina"]) if right["pagina"] is not None else "..."
    pdf.set_xy(x_right, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_r, border=1, align="C")

    # Avanzar Y para la siguiente fila
    pdf.set_y(y + row_h)

# CONFIGURACIÓN DE GRID FLEXIBLE
COLS = 3
PAGE_WIDTH = 210
MARGIN = 10
COLUMN_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / COLS
Y_START = 30
Y_LIMIT = 280

x_positions = [MARGIN + i * COLUMN_WIDTH for i in range(COLS)]
pdf.provincia_actual = ""

line_height = 4
ancho_texto = COLUMN_WIDTH - 8

# Array para rastrear la altura Y actual en cada columna
y_actual = [Y_START] * COLS

# GENERAR FICHAS
provincia_anterior = ""
localidad_anterior = ""
current_col = 0
indice_actual = 0

# Calcular números de página para cada provincia
# Simular el bucle de generación para saber cuándo se añade página
pagina_actual = 3  # Primera página de catálogo (1 portada, 2 índice, 3 catálogo)
columna_actual = 0
altura_y = 30
provincia_anterior_temp = ""

for idx, row in df.iterrows():
    provincia = str(row["PROVINCIA"])

    if provincia != provincia_anterior_temp:
        # Nueva provincia: encontrar en indice_provincias y asignar página
        for item in indice_provincias:
            if item["provincia"] == provincia and item["pagina"] is None:
                item["pagina"] = pagina_actual
                break

        # Actualizar valores para nueva provincia
        provincia_anterior_temp = provincia
        pagina_actual += 1
        columna_actual = 0
        altura_y = 30

# Ahora que tenemos los números de página, regenerar el PDF con índice correcto
# Crear un nuevo PDF desde cero
pdf = PDF()
pdf.set_auto_page_break(auto=False)
pdf.set_font("Helvetica", "", 9)
pdf.set_text_color(0, 0, 0)

# Añadir portada a toda la página si existe
if SHOW_PORTADA:
    try:
        pdf.add_page()
        PAGE_W = pdf.w  # ancho de la página en mm
        PAGE_H = pdf.h  # alto de la página en mm
        pdf.image("portada.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
    except Exception as e:
        # Si falla la carga, seguimos sin portada
        print(f"No se pudo cargar portada.jpg: {e}")

# Añadir segunda página con imagen
try:
    pdf.add_page()
    PAGE_W = pdf.w
    PAGE_H = pdf.h
    pdf.image("Segunda-pagina.jpg", x=0, y=0, w=PAGE_W, h=PAGE_H)
except Exception as e:
    print(f"No se pudo cargar Segunda-pagina.jpg: {e}")

# --- REGENERAR PÁGINA DE ÍNDICE CON NÚMEROS CORRECTOS ---
pdf.add_page()
pdf.set_font("Helvetica", "B", 16)
pdf.set_text_color(0, 0, 0)
pdf.cell(0, 12, "PROVINCIAS DE ESPAÑA Y SUS CAPITALES", ln=True, align="C")
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 10, "PROVINCES OF SPAIN AND THEIR CAPITALS", ln=True, align="C")
pdf.ln(8)

# Tabla de índice en dos columnas paralelas: PROVINCIAS | CAPITALES | Pág.
usable_width = PAGE_WIDTH - 2 * MARGIN
separation = 10
table_width = (usable_width - separation) / 2
col_widths = [table_width * 0.40, table_width * 0.45, table_width * 0.15]
x_left = MARGIN
x_right = MARGIN + table_width + separation
row_h = 8

# Preparar listas izquierda/derecha
n = len(indice_provincias)
mid = (n + 1) // 2
left_items = indice_provincias[:mid]
right_items = indice_provincias[mid:]

# Igualar longitudes para evitar filas vacías al final
while len(left_items) < len(right_items):
    left_items.append({"provincia": "", "capital": "", "pagina": None})
while len(right_items) < len(left_items):
    right_items.append({"provincia": "", "capital": "", "pagina": None})

# Encabezados para ambas tablas EN POSICIÓN NATURAL
pdf.set_font("Helvetica", "B", 10)
y_header_start = pdf.get_y()

# Imprimir encabezado izquierdo
pdf.set_xy(x_left, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Imprimir encabezado derecho (MISMA ALTURA)
pdf.set_xy(x_right, y_header_start)
pdf.cell(col_widths[0], row_h, "PROVINCIAS", border=1, align="C")
pdf.cell(col_widths[1], row_h, "CAPITALES", border=1, align="C")
pdf.cell(col_widths[2], row_h, "Pág.", border=1, align="C")

# Avanzar posición después de la cabecera
y_first_row = y_header_start + row_h
pdf.set_y(y_first_row)

# Imprimir filas en paralelo (sin filas vacías al final)
pdf.set_font("Helvetica", "", 10)
for i in range(len(left_items)):
    left = left_items[i]
    right = (
        right_items[i]
        if i < len(right_items)
        else {"provincia": "", "capital": "", "pagina": None}
    )

    prov_l = left["provincia"]
    prov_r = right["provincia"]

    # Saltar si ambas provincias están vacías
    if not prov_l and not prov_r:
        continue

    y = pdf.get_y()

    # Imprimir fila izquierda
    capital_l = left["capital"]
    page_l = str(left["pagina"]) if left["pagina"] is not None else "..."
    pdf.set_xy(x_left, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_l.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_l, border=1, align="C")

    # Imprimir fila derecha (MISMO Y que la izquierda)
    capital_r = right["capital"]
    page_r = str(right["pagina"]) if right["pagina"] is not None else "..."
    pdf.set_xy(x_right, y)
    pdf.cell(
        col_widths[0],
        row_h,
        prov_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(
        col_widths[1],
        row_h,
        capital_r.encode("latin-1", "ignore").decode("latin-1"),
        border=1,
        align="L",
    )
    pdf.cell(col_widths[2], row_h, page_r, border=1, align="C")

    # Avanzar Y para la siguiente fila
    pdf.set_y(y + row_h)

# --- Eliminar página "plantilla" en blanco justo después del índice (si existe) ---
try:
    cur_page = pdf.page_no()  # página actual (la del índice)
    pages_obj = getattr(pdf, "pages", None)
    if pages_obj and len(pages_obj) > cur_page:
        next_page = pages_obj[cur_page]  # páginas en pdf.pages son 0-based
        if isinstance(next_page, str) and not next_page.strip():
            # quitar la página vacía
            pages_obj.pop(cur_page)
except Exception:
    # si algo falla no interrumpimos la generación
    pass

# VOLVER A GENERAR CATÁLOGO (reset de variables)
x_positions = [MARGIN + i * COLUMN_WIDTH for i in range(COLS)]
pdf.provincia_actual = ""
y_actual = [Y_START] * COLS

provincia_anterior = ""
localidad_anterior = ""
current_col = 0

# GENERAR FICHAS (segunda pasada)
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

    # ---- REGISTRAR HOTEL CON SU PÁGINA ACTUAL ----
    if hotel_name and hotel_name not in hotel_pages:
        hotel_pages[hotel_name] = pdf.page_no()

    # CONSTRUIR TEXTO CON FORMATO (negrita para nombre del hotel)
    # Línea 1: Clasificación + habitaciones
    linea1 = (
        f"{row['CLASIFICACION HOTEL']} · {str(row['NRO. HABITACIONES'])} Habitaciones"
    )
    # Línea 2: Nombre del hotel
    linea2 = f"{row['NOMBRE DE EMPRESA']}"
    linea3 = f"{row['DIRECCION']}".title()
    linea4 = f"{row['CP']} {row['LOCALIDAD']}".title()
    linea5 = f"Tel: {str(row['TELEFONO1'])}".title()
    linea6 = f"{row['SITIO WEB']}".lower()

    # Convertir a latin-1
    linea1 = linea1.encode("latin-1", "ignore").decode("latin-1")
    linea2 = linea2.encode("latin-1", "ignore").decode("latin-1")
    linea3 = linea3.encode("latin-1", "ignore").decode("latin-1")
    linea4 = linea4.encode("latin-1", "ignore").decode("latin-1")
    linea5 = linea5.encode("latin-1", "ignore").decode("latin-1")
    linea6 = linea6.encode("latin-1", "ignore").decode("latin-1")

    lineas_hotel = [linea1, linea2, linea3, linea4, linea5, linea6]

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

    # Línea 1: NORMAL (clasificación y habitaciones)
    pdf.set_font("Helvetica", "", 8)
    pdf.multi_cell(ancho_texto, line_height, linea1, border=0, align="L")

    # Línea 2: NEGRITA (nombre del hotel)
    pdf.set_x(x + 2)
    pdf.set_font("Helvetica", "B", 9)
    pdf.multi_cell(ancho_texto, line_height, linea2, border=0, align="L")

    # Líneas 3-6: NORMAL
    pdf.set_font("Helvetica", "", 8)
    pdf.set_x(x + 2)
    pdf.multi_cell(ancho_texto, line_height, linea3, border=0, align="L")
    pdf.set_x(x + 2)
    pdf.multi_cell(ancho_texto, line_height, linea4, border=0, align="L")
    pdf.set_x(x + 2)
    pdf.multi_cell(ancho_texto, line_height, linea5, border=0, align="L")
    pdf.set_x(x + 2)
    pdf.multi_cell(ancho_texto, line_height, linea6, border=0, align="L")

    # Actualizar y_actual[current_col] con la nueva posición Y después del hotel
    y_actual[current_col] = pdf.get_y() + 2  # pequeño espaciado entre hoteles

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

# Configuración: 5 columnas verticales
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


# ---- DISTRIBUIR HOTELES EN COLUMNAS VERTICALES ----
n_hotels = len(hoteles_lista)
hotels_per_col = (n_hotels + COLS_INDEX - 1) // COLS_INDEX  # Redondear hacia arriba

# Crear estructura de columnas (cada columna es una lista)
columnas = [[] for _ in range(COLS_INDEX)]
for i, hotel in enumerate(hoteles_lista):
    col_idx = i // hotels_per_col  # Determinar columna
    if col_idx >= COLS_INDEX:
        col_idx = COLS_INDEX - 1
    columnas[col_idx].append(hotel)

# ---- IMPRIMIR COLUMNAS VERTICALES ----
pdf.set_font("Helvetica", "", 6.5)  # Tipografía pequeña (equivalente a 6pt)
pdf.set_text_color(0, 0, 0)

x_cols = [MARGIN + i * col_width_index for i in range(COLS_INDEX)]
y_cols = [y_start_index] * COLS_INDEX  # Altura actual de cada columna

# Iterar columna por columna (llenado vertical)
page_index = 0
while any(y_cols[i] < y_limit_index and i < COLS_INDEX for i in range(COLS_INDEX)):

    # Verificar si alguna columna necesita nueva página
    need_new_page = False
    for col_idx in range(COLS_INDEX):
        # Contar cuántos hoteles faltan en esta columna
        hotels_done_in_col = (
            sum(1 for y in [y_cols[col_idx]] if y > y_start_index)
            if page_index == 0
            else 0
        )

        # Si y > límite y aún hay hoteles, necesita nueva página
        if y_cols[col_idx] > y_limit_index and any(
            len(columnas[col_idx]) > 0 for col_idx in range(COLS_INDEX)
        ):
            need_new_page = True
            break

    if need_new_page:
        # Nueva página completa
        pdf.add_page()
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
        pdf.set_font("Helvetica", "", 6.5)

        # Reset de alturas
        y_cols = [pdf.get_y()] * COLS_INDEX
        y_start_index = pdf.get_y()
        page_index += 1

    # Llenar 1 fila en cada columna (de izquierda a derecha)
    filled_in_row = False

    for col_idx in range(COLS_INDEX):
        # Encontrar siguiente hotel no impreso en esta columna
        hotel_offset = page_index * (n_hotels // COLS_INDEX + 1)  # Aproximación

        # Búsqueda simple: iterar columnas para encontrar elemento en orden
        # En realidad, ya está distribuido en columnas[], solo falta rastrearlo
        # Usar índice global para este propósito
        pass

    # Mejor aproximación: usar índice lineal y mapear a columnas
    page_index += 1
    if page_index > 2:
        break

# ---- APROXIMACIÓN ALTERNATIVA MÁS SIMPLE: FILL-BY-ROWS (pero en vertical) ----
# Resetear e implementar más directo

# Recalcular desde cero con método más limpio
pdf.set_font("Helvetica", "", 6.5)
pdf.set_text_color(0, 0, 0)

# Ir a la posición después del título
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

pdf.output(PDF_FILE)
print("PDF generado con índice alfabético de 5 columnas verticales:", PDF_FILE)
