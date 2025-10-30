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
    if col_name not in row.index:
        return None

    value = row[col_name]

    # Si hay duplicados ('col', 'col_1'), row[col_name] puede devolver una Serie
    if isinstance(value, pd.Series):
        value = value.iloc[0]  # Tomar el primer valor

    if pd.notna(value):
        return str(value).strip()
    return None


def evaluar_criterios(row, bank_name):
    """
    Evalúa los criterios para una aplicación (fila) basado en la lógica
    del script 1_generador_madurez_y_reportes.py.
    """
    resultados = {}

    # --- 1. OBSOLESCENCIA ---
    # Nota: Los nombres de columna ya están limpios por load_database.py
    valor_obs = get_value_from_row(row, "nivel_de_obsolescencia")
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
    escalabilidad_map = {"SI": "Cumple", "NO": "No Cumple"}
    # Usamos el nombre de columna acortado de la BD
    valor_esc_raw = get_value_from_row(row, "tiene_alta_disponibilidad")
    valor_esc = valor_esc_raw.upper() if valor_esc_raw else None
    resultados["escalabilidad"] = escalabilidad_map.get(valor_esc, "")

    # --- 3. ACOPLE ---
    resultados["acople"] = "Parcialmente"

    # --- 4. ESTABILIDAD ---
    estabilidad_map = {"NO": "Cumple", "SI": "No Cumple"}
    # Usamos el nombre de columna acortado de la BD
    valor_estab_raw = get_value_from_row(
        row, "ha_presentado_caidas_o_degradacion_del_servicio_en_los_ultimo"
    )
    valor_estab = valor_estab_raw.upper() if valor_estab_raw else None
    resultados["estabilidad"] = estabilidad_map.get(valor_estab, "")

    # --- 5. AGILIDAD ---
    devops_raw = get_value_from_row(row, "devops")
    despliegue_raw = get_value_from_row(row, "despliegue_a_pdn_automatizado")

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
    extensibilidad_map = {
        "Regional": "Cumple",
        "Global": "Cumple",
        "Local": "No Cumple",
    }
    if resultados.get("obsolescencia") == "No Cumple":
        resultados["extensibilidad"] = "No Cumple"
    else:
        valor_ext = get_value_from_row(row, "bns")  # Asumiendo que 'bns' es la columna
        resultados["extensibilidad"] = extensibilidad_map.get(
            valor_ext.title() if valor_ext else None, ""
        )

    # --- 7. SEGURIDAD ---
    valor_seg_raw = get_value_from_row(row, "seguridad")
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
    resultados["cobertura"] = ""  # Lógica no definida en la referencia

    # --- 9. UX ---
    ux_map = {"SI": "Cumple", "NO": "No Cumple"}
    valor_ux_raw = get_value_from_row(row, "ux")
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


# --- FUNCIÓN PRINCIPAL DEL MÓDULO ---


def generate_slide_for_txt(
    app_list_filename,
    df_buyer,
    choices_buyer,
    df_bought,
    choices_bought,
):
    """
    Procesa un solo archivo .txt y genera un .pptx.
    Adaptado de 'process_list_file' de 1_generador_madurez_y_reportes.py
    """

    with open(app_list_filename, "r", encoding="utf-8") as f:
        lines_to_process = [line.strip() for line in f if line.strip()]

    print(
        f"    ▶️  Procesando {len(lines_to_process)} líneas de '{os.path.basename(app_list_filename)}'..."
    )

    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.867), Cm(19.05)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    x_positions = {}
    current_x = Cm(1.5)
    for col_name in COLUMN_ORDER:
        x_positions[col_name] = current_x
        current_x += COLUMN_WIDTHS[col_name]
    draw_main_header(slide, x_positions, Cm(1.0))
    y_pos = Cm(1.0) + ROW_HEIGHT
    map_resultado_a_icono = {
        "Cumple": "si",
        "No Cumple": "no",
        "Parcialmente": "parcial",
    }

    for line in lines_to_process:
        try:
            # Asumimos el formato 'pais', 'banco', 'app_name'
            parts = re.findall(r'"(.*?)"', line)
            if len(parts) != 3:
                print(f"      ⚠️  Línea ignorada (formato incorrecto): '{line}'")
                continue

            country, bank, app_name = [part.strip() for part in parts]
            bank_upper = bank.upper()

            # Lógica neutral
            if "BUYERBANK" in bank_upper:
                target_df, target_choices = df_buyer, choices_buyer
                bank_name_eval = "BuyerBank"
            elif "BOUGHTBANK" in bank_upper:
                target_df, target_choices = df_bought, choices_bought
                bank_name_eval = "BoughtBank"
            else:
                print(
                    f"      ⚠️  Banco no reconocido (use 'BuyerBank' o 'BoughtBank'): '{line}'"
                )
                continue
        except Exception as e:
            print(f"      ❌ Error procesando línea '{line}': {e}")
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
            # Evaluar criterios
            resultados_evaluacion = evaluar_criterios(row, bank_name_eval)

            # Dibujar Iconos (SAS, COTS, CLOUD, REGIONAL)
            val_sas = get_value_from_row(row, "sas")
            if val_sas and normalize_string(val_sas) == "si":
                icon_path = os.path.join(ICONS_FOLDER, "sass.png")
                x_icon = x_positions["sas"] + (COLUMN_WIDTHS["sas"] - ICON_SIZE) / 2
                add_image(slide, icon_path, x_icon, y_icon_pos)

            val_custom = get_value_from_row(row, "nivel_de_customizacion")
            if val_custom:
                valor_customizacion = normalize_string(val_custom)
                if valor_customizacion in ["cots", "cots con observacion"]:
                    icon_path = os.path.join(ICONS_FOLDER, "cots.svg")
                    x_icon = (
                        x_positions["cots"] + (COLUMN_WIDTHS["cots"] - ICON_SIZE) / 2
                    )
                    add_image(slide, icon_path, x_icon, y_icon_pos)

            val_nube = get_value_from_row(row, "nube_vs_onpremise")
            if val_nube and normalize_string(val_nube) == "nube":
                icon_path = os.path.join(ICONS_FOLDER, "cloud.svg")
                x_icon = x_positions["cloud"] + (COLUMN_WIDTHS["cloud"] - ICON_SIZE) / 2
                add_image(slide, icon_path, x_icon, y_icon_pos)

            val_bns_raw = get_value_from_row(row, "bns")
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
            tech_text = get_value_from_row(row, "tecnologia_subyacente")
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
                tf_tec.margin_left = tf_tec.margin_right = tf_tec.margin_top = (
                    tf_tec.margin_bottom
                ) = 0
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

    # Guardar el PPTX
    relative_path = os.path.relpath(app_list_filename, "inputs")
    relative_path_no_ext, _ = os.path.splitext(relative_path)
    output_relative_path = relative_path_no_ext + ".pptx"
    output_path = os.path.join(OUTPUT_FOLDER, output_relative_path)
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)

    try:
        prs.save(output_path)
        print(f"    ✅ ¡Éxito! El archivo '{output_path}' ha sido generado.")
    except Exception as e:
        print(f"    ❌ ERROR al guardar '{output_path}': {e}")
