"""
Contains all logic for generating application slides,
adapted from the reference scripts.
"""

import pandas as pd
import re
import os
import sys
from io import BytesIO
from unidecode import unidecode
from thefuzz import process, fuzz
from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPM
import warnings

# --- IMPORTAMOS LA PLANTILLA BASE ---
from .base_slide import create_base_slide

# Ocultar warnings de openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- CONFIGURACIÓN ---
OUTPUT_FOLDER = "outputs"
ICONS_FOLDER = "icons"
APP_COLUMN_NAME = "aplicacion_sistema"

# Parámetros (tomados de 1_generador_madurez_y_reportes.py)
SIMILARITY_THRESHOLD = 0.90
TECH_TRUNCATE_LENGTH = 33
ROW_HEIGHT = Cm(0.62)
TEXTBOX_HEIGHT = Cm(0.48)
APP_TEXTBOX_WIDTH = Cm(4.27)
TECH_TEXTBOX_WIDTH = Cm(4.27)
ICON_SIZE = Cm(0.46)
FONT_SIZE = 8

INDICATOR_ICONS = {"si": "si.svg", "no": "no.svg", "parcial": "na.svg", "na": "na.svg"}
HEADER_LABELS = {
    "aplicaciones": "Aplicaciones",
    "sas": "SAS",
    "cloud": "Cloud",
    "cots": "COTS",
    "regional": "Regional",
    "tecnologia_subyacente": "Tecnología subyacente",
    "obsolescencia": "Obsolescencia",
    "escalabilidad": "Escalabilidad",
    "acople": "Acople",
    "estabilidad": "Estabilidad",
    "extensibilidad": "Extensibilidad",
    "seguridad": "Seguridad",
    "cobertura": "Cobertura",
    "ux": "UX",
    "agilidad": "Agilidad",
}
# Asegurarnos de usar los nombres de columna de la BD (con '_')
COLUMN_ORDER = [
    "aplicaciones",
    "sas",
    "cloud",
    "cots",
    "regional",
    "tecnologia_subyacente",
    "obsolescencia",
    "escalabilidad",
    "acople",
    "estabilidad",
    "extensibilidad",
    "seguridad",
    "cobertura",
    "ux",
    "agilidad",
]
COLUMN_WIDTHS = {
    "aplicaciones": APP_TEXTBOX_WIDTH,
    "sas": Cm(0.6),
    "cloud": Cm(0.6),
    "cots": Cm(0.6),
    "regional": Cm(0.6),
    "tecnologia_subyacente": TECH_TEXTBOX_WIDTH,
    "obsolescencia": Cm(2.0),
    "escalabilidad": Cm(2.0),
    "acople": Cm(2.0),
    "estabilidad": Cm(2.0),
    "extensibilidad": Cm(2.0),
    "seguridad": Cm(2.0),
    "cobertura": Cm(2.0),
    "ux": Cm(2.0),
    "agilidad": Cm(2.0),
}


# --- FUNCIONES DE LÓGICA (Adaptadas de la referencia) ---


def normalize_string(text):
    """(Función de masters/excel_loader.py)"""
    if not isinstance(text, str):
        return ""
    subscript_map = str.maketrans("₀₁₂₃₄₅₆₇₈₉", " " * 10)
    clean_text = text.translate(subscript_map)
    clean_text = unidecode(clean_text).lower()
    clean_text = re.sub(r"\s*\([^)]*\)\s*", " ", clean_text)
    noise_words = ["incluida en venta", "tsa", "no tsa"]
    for word in noise_words:
        clean_text = clean_text.replace(word, "")
    clean_text = re.sub(r"[^a-z0-9\s-]", "", clean_text)
    clean_text = clean_text.strip()
    return re.sub(r"\s+", " ", clean_text)


def get_value_from_row(row, col_name):
    """
    Obtiene un valor único de una fila (un pd.Series),
    manejando columnas duplicadas (ej. 'banco_1').
    """
    if not col_name or col_name not in row.index:
        return None

    value = row[col_name]

    # Si hay duplicados ('col', 'col_1'), row[col_name] puede devolver una Serie
    if isinstance(value, pd.Series):
        value = value.iloc[0]  # Tomar el primer valor

    if pd.notna(value):
        return str(value).strip()
    return None


def evaluar_criterios(row, bank_name, criteria_map):
    """
    Evalúa los criterios para una aplicación (fila) usando el mapeo
    proporcionado por 'config.py'.
    """
    resultados = {}

    # --- 1. OBSOLESCENCIA ---
    col_obs = criteria_map.get("obsolescencia")
    valor_obs = get_value_from_row(row, col_obs)
    if valor_obs:
        valor_lower = valor_obs.lower()
        if bank_name == "BuyerBank":
            if "vigente" in valor_lower:
                resultados["obsolescencia"] = "Cumple"
            else:
                resultados["obsolescencia"] = "No Cumple"
        elif bank_name == "BoughtBank":
            if "no obsoleto" in valor_lower:
                resultados["obsolescencia"] = "Cumple"
            elif "obsoleto" in valor_lower:
                resultados["obsolescencia"] = "No Cumple"
            else:
                resultados["obsolescencia"] = ""
    else:
        resultados["obsolescencia"] = ""

    # --- 2. ESCALABILIDAD ---
    col_esc = criteria_map.get("escalabilidad")
    escalabilidad_map = {"SI": "Cumple", "NO": "No Cumple"}
    valor_esc_raw = get_value_from_row(row, col_esc)
    valor_esc = valor_esc_raw.upper() if valor_esc_raw else None
    resultados["escalabilidad"] = escalabilidad_map.get(valor_esc, "")

    # --- 3. ACOPLE ---
    if criteria_map.get("acople") is None:
        resultados["acople"] = "Parcialmente"  # Lógica hardcodeada
    else:
        # (Lógica futura si se mapea a una columna)
        resultados["acople"] = ""

    # --- 4. ESTABILIDAD ---
    col_estab = criteria_map.get("estabilidad")
    estabilidad_map = {"NO": "Cumple", "SI": "No Cumple"}
    valor_estab_raw = get_value_from_row(row, col_estab)
    valor_estab = valor_estab_raw.upper() if valor_estab_raw else None
    resultados["estabilidad"] = estabilidad_map.get(valor_estab, "")

    # --- 5. AGILIDAD ---
    col_agilidad_devops, col_agilidad_despliegue = criteria_map.get(
        "agilidad", (None, None)
    )
    devops_raw = get_value_from_row(row, col_agilidad_devops)
    despliegue_raw = get_value_from_row(row, col_agilidad_despliegue)

    devops = devops_raw.upper() if devops_raw else None
    despliegue = despliegue_raw.upper() if despliegue_raw else None

    if devops == "NO":
        resultados["agilidad"] = "No Cumple"
    elif devops == "SI" and despliegue == "SI":
        resultados["agilidad"] = "Cumple"
    elif devops == "SI":
        resultados["agilidad"] = "Parcialmente"
    else:
        resultados["agilidad"] = ""

    # --- 6. EXTENSIBILIDAD ---
    col_ext = criteria_map.get("extensibilidad")
    extensibilidad_map = {
        "Regional": "Cumple",
        "Global": "Cumple",
        "Local": "No Cumple",
    }
    # Lógica de negocio: depende de la obsolescencia
    if resultados.get("obsolescencia") == "No Cumple":
        resultados["extensibilidad"] = "No Cumple"
    else:
        valor_ext = get_value_from_row(row, col_ext)
        resultados["extensibilidad"] = extensibilidad_map.get(
            valor_ext.title() if valor_ext else None, ""
        )

    # --- 7. SEGURIDAD ---
    col_seg = criteria_map.get("seguridad")
    valor_seg_raw = get_value_from_row(row, col_seg)
    try:
        valor_seg_num = float(valor_seg_raw)
        if valor_seg_num <= 2:
            resultados["seguridad"] = "No Cumple"
        elif valor_seg_num == 3:
            resultados["seguridad"] = "Parcialmente"
        elif valor_seg_num >= 4:
            resultados["seguridad"] = "Cumple"
        else:
            resultados["seguridad"] = ""
    except (ValueError, TypeError):
        resultados["seguridad"] = ""

    # --- 8. COBERTURA ---
    if criteria_map.get("cobertura") is None:
        resultados["cobertura"] = ""  # Lógica hardcodeada
    else:
        resultados["cobertura"] = ""

    # --- 9. UX ---
    col_ux = criteria_map.get("ux")
    ux_map = {"SI": "Cumple", "NO": "No Cumple"}
    valor_ux_raw = get_value_from_row(row, col_ux)
    valor_ux = valor_ux_raw.upper() if valor_ux_raw else None
    resultados["ux"] = ux_map.get(valor_ux, "")

    return resultados


def find_best_match(app_name, choices_dict):
    """(Función de 1_generador_madurez_y_reportes.py)"""
    normalized_name = normalize_string(app_name)
    if not choices_dict:
        return None
    if normalized_name in choices_dict:
        return choices_dict[normalized_name]

    # Usar 'process' de thefuzz para encontrar la mejor coincidencia
    match = process.extractOne(
        normalized_name, choices_dict.keys(), scorer=fuzz.token_set_ratio
    )

    if match and match[1] >= SIMILARITY_THRESHOLD * 100:
        return choices_dict[match[0]]

    return None


# --- FUNCIONES DE POWERPOINT (Adaptadas de la referencia) ---


def add_image(slide, img_path, x, y, height=ICON_SIZE):
    """(Función de 1_generador_madurez_y_reportes.py)"""
    if not os.path.exists(img_path):
        return None
    try:
        width = height
        if img_path.lower().endswith(".svg"):
            drawing = svg2rlg(img_path)
            if not drawing:
                return None
            png_buffer = BytesIO()
            renderPM.drawToFile(drawing, png_buffer, fmt="PNG", dpi=300)
            png_buffer.seek(0)
            return slide.shapes.add_picture(
                png_buffer, x, y, width=width, height=height
            )
        elif img_path.lower().endswith((".png", ".jpg", ".jpeg")):
            return slide.shapes.add_picture(img_path, x, y, width=width, height=height)
    except Exception:
        pass
    return None


def draw_main_header(slide, x_positions, y_pos):
    """(Función de 1_generador_madurez_y_reportes.py)"""
    for col_name in COLUMN_ORDER:
        x, width = x_positions[col_name], COLUMN_WIDTHS[col_name]
        textbox = slide.shapes.add_textbox(x, y_pos, width, ROW_HEIGHT)
        tf = textbox.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = HEADER_LABELS[col_name]
        run.font.bold = True
        run.font.size = Pt(FONT_SIZE)


# --- FUNCIÓN PRINCIPAL DEL MÓDULO (MODIFICADA) ---


def generate_slide_for_subdomain(
    prs,
    subdomain_title,
    app_lines_data,
    df_buyer,
    choices_buyer,
    df_bought,
    choices_bought,
    criteria_map,  # <-- NUEVO PARÁMETRO
):
    """
    Crea UN slide para un subdominio específico, usando la plantilla base.
    """

    print(f"    ▶️  Generando slide para el subdominio: {subdomain_title}...")

    # 1. CREAR EL SLIDE USANDO LA PLANTILLA BASE
    slide = create_base_slide(prs, subdomain_title, "")

    # 2. DEFINIR COORDENADAS DE INICIO PARA EL CONTENIDO
    START_X = Cm(1.54)
    START_Y = Cm(5.22)  # Top del área de contenido

    # 3. DIBUJAR LA TABLA
    x_positions = {}
    current_x = START_X
    for col_name in COLUMN_ORDER:
        x_positions[col_name] = current_x
        current_x += COLUMN_WIDTHS[col_name]

    draw_main_header(slide, x_positions, START_Y)

    y_pos = START_Y + ROW_HEIGHT

    map_resultado_a_icono = {
        "Cumple": "si",
        "No Cumple": "no",
        "Parcialmente": "parcial",
    }

    # Iterar sobre las líneas de app (ya no se lee de un archivo)
    for line_data in app_lines_data:
        try:
            country, bank, app_name = line_data
            bank_upper = bank.upper()

            if "BUYERBANK" in bank_upper:
                target_df, target_choices = df_buyer, choices_buyer
                bank_name_eval = "BuyerBank"
            elif "BOUGHTBANK" in bank_upper:
                target_df, target_choices = df_bought, choices_bought
                bank_name_eval = "BoughtBank"
            else:
                print(
                    f"      ⚠️  Banco no reconocido (use 'BuyerBank' o 'BoughtBank'): '{line_data}'"
                )
                continue
        except Exception as e:
            print(f"      ❌ Error procesando línea '{line_data}': {e}")
            continue

        excel_match_name = find_best_match(app_name, target_choices)
        row_df = (
            target_df[target_df[APP_COLUMN_NAME] == excel_match_name]
            if excel_match_name
            else pd.DataFrame()
        )
        row = row_df.iloc[0] if not row_df.empty else None

        y_text_pos = y_pos + (ROW_HEIGHT - TEXTBOX_HEIGHT) / 2
        y_icon_pos = y_pos + (ROW_HEIGHT - ICON_SIZE) / 2

        x_app, w_app = x_positions["aplicaciones"], COLUMN_WIDTHS["aplicaciones"]
        textbox = slide.shapes.add_textbox(x_app, y_text_pos, w_app, TEXTBOX_HEIGHT)
        p = textbox.text_frame.paragraphs[0]
        p.text = app_name
        p.font.size = Pt(FONT_SIZE)

        if row is not None:
            # Evaluar criterios usando el MAPA
            resultados_evaluacion = evaluar_criterios(row, bank_name_eval, criteria_map)

            # Dibujar Iconos (SAS, COTS, CLOUD, REGIONAL)
            val_sas = get_value_from_row(row, criteria_map.get("icon_sas"))
            if val_sas and normalize_string(val_sas) == "si":
                icon_path = os.path.join(ICONS_FOLDER, "sass.png")
                x_icon = x_positions["sas"] + (COLUMN_WIDTHS["sas"] - ICON_SIZE) / 2
                add_image(slide, icon_path, x_icon, y_icon_pos)

            val_custom = get_value_from_row(row, criteria_map.get("icon_cots"))
            if val_custom:
                valor_customizacion = normalize_string(val_custom)
                if valor_customizacion in ["cots", "cots con observacion"]:
                    icon_path = os.path.join(ICONS_FOLDER, "cots.svg")
                    x_icon = (
                        x_positions["cots"] + (COLUMN_WIDTHS["cots"] - ICON_SIZE) / 2
                    )
                    add_image(slide, icon_path, x_icon, y_icon_pos)

            val_nube = get_value_from_row(row, criteria_map.get("icon_cloud"))
            if val_nube and normalize_string(val_nube) == "nube":
                icon_path = os.path.join(ICONS_FOLDER, "cloud.svg")
                x_icon = x_positions["cloud"] + (COLUMN_WIDTHS["cloud"] - ICON_SIZE) / 2
                add_image(slide, icon_path, x_icon, y_icon_pos)

            val_bns_raw = get_value_from_row(row, criteria_map.get("icon_regional"))
            if val_bns_raw:
                valor_bns_norm = normalize_string(val_bns_raw)
                if "regional" in valor_bns_norm or "global" in valor_bns_norm:
                    icon_path = os.path.join(ICONS_FOLDER, "regional.svg")
                    x_icon = (
                        x_positions["regional"]
                        + (COLUMN_WIDTHS["regional"] - ICON_SIZE) / 2
                    )
                    add_image(slide, icon_path, x_icon, y_icon_pos)

            # Escribir Tecnología
            tech_text = get_value_from_row(row, criteria_map.get("tecnologia"))
            if tech_text:
                if len(tech_text) > TECH_TRUNCATE_LENGTH:
                    tech_text = tech_text[:TECH_TRUNCATE_LENGTH] + "..."
                textbox_tec = slide.shapes.add_textbox(
                    x_positions["tecnologia_subyacente"],
                    y_text_pos,
                    COLUMN_WIDTHS["tecnologia_subyacente"],
                    TEXTBOX_HEIGHT,
                )
                tf_tec = textbox_tec.text_frame
                (
                    tf_tec.margin_left,
                    tf_tec.margin_right,
                    tf_tec.margin_top,
                    tf_tec.margin_bottom,
                ) = 0, 0, 0, 0
                tf_tec.vertical_anchor = MSO_ANCHOR.MIDDLE
                p_tec = tf_tec.paragraphs[0]
                p_tec.text = tech_text
                p_tec.font.size = Pt(FONT_SIZE)

            # Dibujar Iconos de Evaluación
            indicator_cols = [
                c
                for c in COLUMN_ORDER
                if c
                not in [
                    "aplicaciones",
                    "sas",
                    "cloud",
                    "cots",
                    "regional",
                    "tecnologia_subyacente",
                ]
            ]
            for col_name in indicator_cols:
                resultado_evaluado = resultados_evaluacion.get(col_name)
                icono_key = map_resultado_a_icono.get(resultado_evaluado)
                if icono_key and icono_key in INDICATOR_ICONS:
                    icon_path = os.path.join(ICONS_FOLDER, INDICATOR_ICONS[icono_key])
                    x_icon = (
                        x_positions[col_name]
                        + (COLUMN_WIDTHS[col_name] - ICON_SIZE) / 2
                    )
                    add_image(slide, icon_path, x_icon, y_icon_pos)

        y_pos += ROW_HEIGHT
