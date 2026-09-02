"""Genera la cubierta completa (contraportada + lomo + portada) para KDP.

Medidas tomadas del calculador oficial de KDP para:
    tapa blanda | tinta blanco y negro | papel blanco | 6" x 9" | 526 paginas

    Portada completa ....... 341,12 x 234,95 mm
    Portada / contraportada  152,40 x 228,60 mm cada una
    Lomo ................... 29,97 mm
    Zona segura del lomo ... 26,80 x 222,25 mm
    Sangrado ............... 3,17 mm por los cuatro bordes exteriores
    Margen de seguridad .... 3,17 mm desde el corte
    Codigo de barras ....... 50,8 x 30,5 mm libres en la esquina inferior
                             derecha de la contraportada, a 6,35 mm del borde

Si cambia el numero de paginas del interior basta con tocar PAGINAS: el lomo y
el ancho total se recalculan solos.
"""

import base64
import io
import os
import re

from fpdf import FPDF

# ---------------------------------------------------------------------------
# PARAMETROS DEL LIBRO
# ---------------------------------------------------------------------------
PAGINAS = 526
# Debe coincidir EXACTAMENTE con "Tinta y tipo de papel" en KDP: cada opcion
# tiene un grosor de hoja distinto y por tanto un lomo distinto.
PAPEL = "color_estandar"      # negro_blanco | negro_crema | color_estandar | color_premium
SALIDA = f"portada-kdp-{PAPEL}.pdf"
MAPA_SVG = "mapa_espana_qualithotelsbook.svg"

GROSOR_HOJA = {          # pulgadas de grosor por pagina, tabla oficial de KDP
    "negro_blanco": 0.002252,
    "negro_crema": 0.0025,
    "color_estandar": 0.002347,
    "color_premium": 0.002252,
}

MM = 25.4
TRIM_W, TRIM_H = 152.4, 228.6         # 6" x 9"
SANGRADO = 3.175                      # 0.125"
SEGURIDAD = 3.175                     # margen minimo desde el corte
COD_BARRAS_W, COD_BARRAS_H = 50.8, 30.5   # 2" x 1.2"
COD_BARRAS_MARGEN = 6.35                  # 0.25"

LOMO = PAGINAS * GROSOR_HOJA[PAPEL] * MM
PAGE_W = 2 * SANGRADO + 2 * TRIM_W + LOMO
PAGE_H = TRIM_H + 2 * SANGRADO

# Coordenadas horizontales de cada panel
CONTRA_X0 = SANGRADO                  # borde de corte izquierdo de la contraportada
CONTRA_X1 = SANGRADO + TRIM_W
LOMO_X0 = CONTRA_X1
LOMO_X1 = LOMO_X0 + LOMO
PORT_X0 = LOMO_X1
PORT_X1 = PORT_X0 + TRIM_W
Y0, Y1 = SANGRADO, SANGRADO + TRIM_H  # bordes de corte superior e inferior

# ---------------------------------------------------------------------------
# TIPOGRAFIA Y COLOR
# ---------------------------------------------------------------------------
FUENTES = {
    ("sans", ""): "C:/Windows/Fonts/arial.ttf",
    ("sans", "B"): "C:/Windows/Fonts/arialbd.ttf",
    ("serif", ""): "C:/Windows/Fonts/times.ttf",
    ("serif", "B"): "C:/Windows/Fonts/timesbd.ttf",
    ("serif", "I"): "C:/Windows/Fonts/timesi.ttf",
    ("serif", "BI"): "C:/Windows/Fonts/timesbi.ttf",
}

AZUL = (26, 42, 88)        # azul marino de los titulos
GRIS = (60, 60, 60)        # texto corrido de la contraportada
NEGRO = (0, 0, 0)


# ---------------------------------------------------------------------------
# TEXTOS
# ---------------------------------------------------------------------------
MARCA = "QualitHotelsBook"
TITULO_ES = "GUIA DE HOTELES DE ESPAÑA"
TITULO_EN = "SPAIN HOTELS GUIDE"
FECHA = "Septiembre 2026"

CONTRA_ES = [
    ("p", "Contiene el nombre, dirección y datos principales de los hoteles "
          "existentes en España. Para la obtención de los datos se han utilizado "
          "fuentes de dominio público y las páginas web de los propios hoteles."),
    ("h", "VENTAJAS DE ESTA GUÍA:"),
    ("d", ("Seguridad: ", "los establecimientos que aparecen en la misma disponen "
           "del Código de Registro Oficial que acredita su legalidad para el "
           "ejercicio de la actividad hotelera.")),
    ("d", ("Garantía de contacto inmediato: ", "Por teléfono fijo, teléfono móvil, "
           "o a través de la web de los hoteles.")),
]

CONTRA_EN = [
    ("p", "This guide contains the names, addresses, and key information of hotels "
          "located throughout Spain. The information has been compiled using "
          "publicly available sources and the websites of the hotels themselves."),
    ("h", "ADVANTAGES OF THIS GUIDE:"),
    ("d", ("Security: ", "All establishments included in this guide have an Official "
           "Registration Code, certifying that they are legally authorized to "
           "operate as hotel establishments.")),
    ("d", ("Immediate Contact Guarantee: ", "Hotels can be contacted directly via "
           "landline, mobile phone, or through their website.")),
]

EDITORIAL = [
    "Editada por QUALIT ASESORES S.L.",
    "Calle Pau Claris, 77 – 08010 Barcelona",
    "E-mail: qualithotelsbook@gmail.com",
    "Teléfono: +34 683.323.608",
    "Web: www.qualithotelsbook.com",
    "",
    "Prohibida su reproducción parcial o total, sin autorización escrita.",
    "Reproduction in whole or in part is prohibited without written authorization.",
]


# ---------------------------------------------------------------------------
# MAPA: el SVG lleva dentro un PNG en base64; lo extraemos a un fichero temporal
# ---------------------------------------------------------------------------
def extraer_mapa(ruta_svg):
    """Devuelve (ruta_png, ancho_px, alto_px) del mapa incrustado en el SVG."""
    from PIL import Image

    svg = io.open(ruta_svg, encoding="utf-8", errors="replace").read()
    m = re.search(r'(?:xlink:)?href="data:image/(\w+);base64,([^"]+)"', svg)
    if not m:
        raise RuntimeError(f"No hay imagen incrustada en {ruta_svg}")
    img = Image.open(io.BytesIO(base64.b64decode(m.group(2))))
    if img.mode in ("RGBA", "LA", "P"):
        fondo = Image.new("RGB", img.size, (255, 255, 255))
        img = img.convert("RGBA")
        fondo.paste(img, mask=img.split()[-1])
        img = fondo
    destino = "_mapa_portada.png"
    img.save(destino)
    # El mapa tiene una franja clara en la parte superior (mar). Medimos donde
    # empieza el dibujo de verdad para poder alinearlo como en el original.
    import numpy as np
    gris = np.asarray(img.convert("L"))
    ys = np.where((gris < 200).any(axis=1))[0]
    offset = ys.min() / gris.shape[0]
    return destino, img.size[0], img.size[1], offset


# ---------------------------------------------------------------------------
# MEDIDAS DE LA PORTADA, TOMADAS DEL DISEÑO ORIGINAL
# ---------------------------------------------------------------------------
# `cap_top`: mm desde el borde de corte superior hasta lo alto de las
#            mayusculas.  `ancho`: anchura de la linea, en mm.
CAP_TIMES = 0.662          # altura de mayuscula de Times New Roman, en ems

PORTADA = {
    "titulo":    dict(texto=MARCA, estilo="B",  cap_top=11.51, ancho=107.86,
                      sufijo="®"),
    "titulo_es": dict(texto=TITULO_ES, estilo="B",  cap_top=34.92, ancho=122.57),
    "titulo_en": dict(texto=TITULO_EN, estilo="B",  cap_top=50.98, ancho=116.55),
    "fecha":     dict(texto=FECHA, estilo="BI", cap_top=66.81, ancho=43.90),
    "mapa":      dict(top=84.25, ancho=128.40),
}

# El titulo en ingles es mas corto pero en el original ocupa casi lo mismo:
# lleva las letras separadas.  Para reproducirlo se fija su cuerpo igual al del
# titulo en español y se reparte la diferencia como espaciado entre caracteres.
IGUALAR_CUERPO = {"titulo_en": "titulo_es"}


def _ancho(pdf, texto, estilo, size, tracking=0.0, sufijo=None):
    """Anchura visible de la linea, espaciado y simbolo ® incluidos."""
    pdf.set_font("serif", estilo, size)
    w = pdf.get_string_width(texto) + tracking * max(len(texto) - 1, 0)
    if sufijo:
        pdf.set_font("serif", estilo, size * 0.40)
        w += size * 0.06 + pdf.get_string_width(sufijo)
    return w


def tamano_por_ancho(pdf, texto, estilo, objetivo, sufijo=None):
    """Cuerpo (en puntos) con el que la linea mide exactamente `objetivo` mm."""
    lo, hi = 4.0, 90.0
    for _ in range(60):
        mid = (lo + hi) / 2
        if _ancho(pdf, texto, estilo, mid, sufijo=sufijo) < objetivo:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def linea_portada(pdf, cx, cap_top, texto, estilo, size, ancho_objetivo,
                  tracking=0.0, sufijo=None):
    """Escribe una linea centrada en `cx` con la parte alta de las mayusculas
    exactamente en `cap_top`."""
    base = cap_top + CAP_TIMES * size / 72 * MM        # linea de base
    w = _ancho(pdf, texto, estilo, size, tracking, sufijo)
    x = cx - w / 2
    pdf.set_text_color(*AZUL)
    if tracking:
        # OJO: set_char_spacing espera PUNTOS, no las unidades del documento.
        pdf.set_char_spacing(tracking / MM * 72)
    pdf.set_font("serif", estilo, size)
    pdf.text(x, base, texto)
    if tracking:
        pdf.set_char_spacing(0)
    if sufijo:
        x_sf = x + pdf.get_string_width(texto) + tracking * max(len(texto) - 1, 0) + size * 0.06
        pdf.set_font("serif", estilo, size * 0.40)
        pdf.text(x_sf, cap_top + CAP_TIMES * size * 0.40 / 72 * MM + size * 0.06, sufijo)


# ---------------------------------------------------------------------------
# CUBIERTA
# ---------------------------------------------------------------------------
class Cubierta(FPDF):
    def __init__(self):
        # OJO: con orientation="L" fpdf2 intercambia ancho y alto del formato.
        # Como ya le damos la medida final (mas ancha que alta), va en "P".
        super().__init__(orientation="P", unit="mm", format=(PAGE_W, PAGE_H))
        self.set_auto_page_break(False)
        self.set_margins(0, 0, 0)
        for (familia, estilo), ruta in FUENTES.items():
            self.add_font(familia, estilo, ruta)


def texto_centrado(pdf, x0, ancho, y, txt, familia, estilo, size, color, alto=None):
    """Escribe una linea centrada en [x0, x0+ancho] y devuelve la Y siguiente."""
    alto = alto or size * 0.42
    pdf.set_font(familia, estilo, size)
    pdf.set_text_color(*color)
    pdf.set_xy(x0, y)
    pdf.cell(ancho, alto, txt, align="C")
    return y + alto


def bloque_contraportada(pdf, x0, ancho, y, titulo, marca, bloques):
    """Dibuja uno de los dos bloques (español / inglés) de la contraportada."""
    pdf.set_font("sans", "B", 17)
    pdf.set_text_color(*NEGRO)
    pdf.set_xy(x0, y)
    pdf.cell(ancho, 7.5, titulo, align="C")
    y += 8.6

    pdf.set_font("sans", "B", 14.5)
    pdf.set_xy(x0, y)
    pdf.cell(ancho, 6.5, marca, align="C")
    y += 11

    for tipo, contenido in bloques:
        if tipo == "p":
            pdf.set_font("sans", "", 9)
            pdf.set_text_color(*GRIS)
            pdf.set_xy(x0, y)
            pdf.multi_cell(ancho, 4.8, contenido, align="J")
            y = pdf.get_y() + 5.0
        elif tipo == "h":
            pdf.set_font("sans", "B", 9.2)
            pdf.set_text_color(*NEGRO)
            pdf.set_xy(x0, y)
            pdf.cell(ancho, 4.8, contenido, align="L")
            y += 7.5
        else:
            # La etiqueta va en negrita y el resto normal DENTRO del mismo
            # multi_cell: asi el ajuste de linea respeta el ancho de columna.
            # Con pdf.write() el texto se desbordaba hacia el lomo.
            etiqueta, resto = contenido
            pdf.set_font("sans", "", 9)
            pdf.set_text_color(*GRIS)
            pdf.set_xy(x0, y)
            pdf.multi_cell(ancho, 4.8, f"**{etiqueta.strip()}** {resto}",
                           align="J", markdown=True)
            y = pdf.get_y() + 4.6
    return y


def construir():
    mapa, mapa_w, mapa_h, mapa_offset = extraer_mapa(MAPA_SVG)
    pdf = Cubierta()
    pdf.add_page()

    # Fondo blanco en TODA la hoja, sangrado incluido
    pdf.set_fill_color(255, 255, 255)
    pdf.rect(0, 0, PAGE_W, PAGE_H, "F")

    # ---------------- CONTRAPORTADA ----------------
    marg = 14.0
    cx0 = CONTRA_X0 + marg
    canc = TRIM_W - 2 * marg
    y = Y0 + 18

    y = bloque_contraportada(pdf, cx0, canc, y, TITULO_ES, MARCA, CONTRA_ES)

    # separador fino centrado
    y += 5.0
    pdf.set_draw_color(120, 120, 120)
    pdf.set_line_width(0.3)
    pdf.line(cx0 + canc / 2 - 24, y, cx0 + canc / 2 + 24, y)
    y += 10.0

    y = bloque_contraportada(pdf, cx0, canc, y, TITULO_EN, MARCA, CONTRA_EN)

    # Datos editoriales, abajo a la izquierda, sin invadir el codigo de barras
    ancho_editorial = (CONTRA_X1 - COD_BARRAS_MARGEN - COD_BARRAS_W) - cx0 - 4
    # Las lineas largas se parten en dos, asi que hay que MEDIR el alto real
    # antes de anclar el bloque al pie: si no, se sale del margen de seguridad.
    ALTO_LINEA = 3.9
    pdf.set_font("sans", "", 7.6)
    n_lineas = 0
    for linea in EDITORIAL:
        if not linea:
            n_lineas += 1
            continue
        n_lineas += len(pdf.multi_cell(ancho_editorial, ALTO_LINEA, linea,
                                       align="L", dry_run=True, output="LINES"))
    y_edit = Y1 - SEGURIDAD - 3.0 - n_lineas * ALTO_LINEA
    pdf.set_text_color(*GRIS)
    for linea in EDITORIAL:
        pdf.set_xy(cx0, y_edit)
        if linea:
            pdf.multi_cell(ancho_editorial, ALTO_LINEA, linea, align="L")
            y_edit = pdf.get_y()
        else:
            y_edit += ALTO_LINEA

    # ---------------- LOMO ----------------
    lomo_txt = f"{MARCA}®  —  {TITULO_ES}  —  {TITULO_EN}  —  {FECHA}"
    cx, cy = (LOMO_X0 + LOMO_X1) / 2, PAGE_H / 2
    largo_util = TRIM_H - 2 * SEGURIDAD          # 222.25 mm
    with pdf.rotation(-90, cx, cy):
        pdf.set_font("serif", "B", 12)
        pdf.set_text_color(*AZUL)
        pdf.set_xy(cx - largo_util / 2, cy - 4)
        pdf.cell(largo_util, 8, lomo_txt, align="C")

    # ---------------- PORTADA ----------------
    # Todas las medidas salen de medir el diseño original pixel a pixel:
    # anchura de cada linea de texto y altura a la que empieza, en mm desde
    # la esquina superior izquierda del RECORTE del panel de portada.
    cxp = PORT_X0 + TRIM_W / 2          # eje central del panel

    cuerpos = {}
    for clave in ("titulo", "titulo_es", "titulo_en", "fecha"):
        m = PORTADA[clave]
        tracking = 0.0
        if clave in IGUALAR_CUERPO:
            # Mismo cuerpo que la linea de referencia; la anchura que falta se
            # consigue separando las letras, como en el diseño original.
            size = cuerpos[IGUALAR_CUERPO[clave]]
            pdf.set_font("serif", m["estilo"], size)
            natural = pdf.get_string_width(m["texto"])
            tracking = (m["ancho"] - natural) / max(len(m["texto"]) - 1, 1)
        else:
            size = tamano_por_ancho(pdf, m["texto"], m["estilo"], m["ancho"],
                                    sufijo=m.get("sufijo"))
        cuerpos[clave] = size
        linea_portada(pdf, cxp, Y0 + m["cap_top"], m["texto"], m["estilo"],
                      size, ancho_objetivo=m["ancho"], tracking=tracking,
                      sufijo=m.get("sufijo"))
    print("  cuerpos: " + ", ".join(f"{k} {v:.1f}pt" for k, v in cuerpos.items()))

    # Mapa: mismo ancho relativo y mismo hueco respecto al texto que el original
    m = PORTADA["mapa"]
    mapa_ancho = m["ancho"]
    mapa_alto = mapa_ancho * mapa_h / mapa_w
    # `top` es donde debe empezar el DIBUJO, no la caja de la imagen
    y_mapa = Y0 + m["top"] - mapa_offset * mapa_alto
    pdf.image(mapa, x=cxp - mapa_ancho / 2, y=y_mapa, w=mapa_ancho, h=mapa_alto)

    pdf.output(SALIDA)
    os.remove(mapa)

    ppp = mapa_w / (mapa_ancho / MM)
    print(f"Cubierta generada: {SALIDA}")
    print(f"  {PAGE_W:.2f} x {PAGE_H:.2f} mm  ({PAGE_W/MM:.3f}\" x {PAGE_H/MM:.3f}\")")
    print(f"  papel {PAPEL}, {PAGINAS} paginas -> lomo {LOMO:.2f} mm")
    print(f"  lomo de x={LOMO_X0:.2f} a x={LOMO_X1:.2f} mm")
    print(f"  mapa colocado a {mapa_ancho:.1f} x {mapa_alto:.1f} mm -> {ppp:.0f} ppp")


if __name__ == "__main__":
    construir()
