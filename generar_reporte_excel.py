"""
Script para generar un reporte de cumplimiento en Excel.

Este script calcula el cumplimiento de criterios (Obsolescencia, etc.)
basándose ÚNICAMENTE en las aplicaciones que fueron marcadas como
'mostrar_en_arquitectura_target' = 'Si' en la base de datos
(es decir, las que se encontraron durante la ejecución de main.py).

El reporte final se ordena según el 'CUSTOM_SORT_ORDER' y
se divide en 8 hojas separadas (4 Resúmenes, 4 Datos Raw)
basado en la combinación de Banco y Región.

Agrupa los subdominios partidos (ej. _1, _2) en una sola fila.
"""

import pandas as pd
import os
import sys
import re
from collections import defaultdict
from sqlalchemy import create_engine

# --- Importar la lógica central del proyecto ---
try:
    from masters.db_loader import load_data_from_db
    from masters.slide_generator import (
        evaluar_criterios,
        normalize_string,
        find_best_match,
        COUNTRY_ICONS,  # Importamos el mapeo de países
    )
    from config import CRITERIA_DB_MAP, CUSTOM_SORT_ORDER
    from main import APP_COLUMN_NAME_DB  # Reusamos la constante
except ImportError as e:
    print(f"Error de importación: {e}")
    print("Asegúrate de que este script esté en la raíz del proyecto.")
    sys.exit(1)

# --- CONFIGURACIÓN ---
INPUT_FOLDER = "inputs"
OUTPUT_FOLDER = "outputs"
EXCEL_REPORT_FILE = os.path.join(OUTPUT_FOLDER, "Reporte_Cumplimiento_Target.xlsx")

# Los criterios que se están evaluando (de config.py y slide_generator.py)
CRITERIA_LIST = [
    "obsolescencia",
    "escalabilidad",
    "acople",
    "estabilidad",
    "extensibilidad",
    "seguridad",
]

# Mapeo de puntajes para el porcentaje
COMPLIANCE_MAP = {
    "Cumple": 1.0,
    "Parcialmente": 0.5,
    "No Cumple": 0.0,
    "N/A": None,  # No cuenta para el promedio
    "": None,  # No cuenta para el promedio
}


# --- LÓGICA DE REGIÓN ---
def get_region(country_str):
    """Clasifica un país como 'Colombia' o 'CAM'."""
    if "Colombia (CO)" in country_str:
        return "Colombia"
    # Asume que todos los demás países definidos son CAM
    elif country_str in COUNTRY_ICONS:
        return "CAM"
    else:
        return "Otro"  # Fallback por si acaso


def get_subdomain_sort_key(domain, subdomain_filename):
    """
    Obtiene la clave numérica de ordenamiento para un subdominio
    basado en el CUSTOM_SORT_ORDER de config.py.
    """
    # 1. Quitar el prefijo del dominio (ej. 'canales_')
    sub_base = subdomain_filename.replace(f"{domain}_", "")
    # 2. Quitar el sufijo numérico (ej. '_1', '_2')
    sub_base = re.sub(r"_\d+$", "", sub_base)

    order_list = CUSTOM_SORT_ORDER.get(domain)

    if order_list and sub_base in order_list:
        return order_list.index(sub_base)

    # Si no se encuentra, poner al final
    return 999


def generar_reporte():
    """
    Función principal para generar el reporte de Excel.
    """
    print("Iniciando generación de reporte Excel de cumplimiento...")

    all_app_evaluations = []

    # --- 1. Cargar datos de la BD ---
    print("Conectando a la base de datos...")
    (df_buyer, df_bought, db_engine) = load_data_from_db()

    if df_buyer is None or df_bought is None:
        print("Error: No se pudieron cargar los datos de la base de datos. Abortando.")
        return

    # --- 2. Crear diccionarios de búsqueda (choices) ---
    choices_buyer = {
        normalize_string(name): name
        for name in df_buyer[APP_COLUMN_NAME_DB].dropna().unique()
    }
    choices_bought = {
        normalize_string(name): name
        for name in df_bought[APP_COLUMN_NAME_DB].dropna().unique()
    }

    # --- 3. Escanear archivos .txt para obtener el contexto (Dominio/Subdominio) ---
    print(f"Escaneando '{INPUT_FOLDER}' en busca de dominios...")
    domain_folders = [
        f
        for f in os.listdir(INPUT_FOLDER)
        if os.path.isdir(os.path.join(INPUT_FOLDER, f))
    ]

    if not domain_folders:
        print(f"Aviso: No se encontraron carpetas de dominio en '{INPUT_FOLDER}'.")
        return

    print(f"Procesando {len(domain_folders)} dominios...")

    for domain_name in domain_folders:
        domain_path = os.path.join(INPUT_FOLDER, domain_name)
        domain_key = domain_name.lower()

        all_txt_files_in_folder = {
            f for f in os.listdir(domain_path) if f.endswith(".txt")
        }

        # Obtener lista ordenada de archivos (lógica de main.py)
        files_to_process_in_order = []
        processed_files_set = set()
        ordered_base_names = CUSTOM_SORT_ORDER.get(domain_key)

        if ordered_base_names:
            for base_name in ordered_base_names:
                base_filename_prefix = f"{domain_key}_{base_name}"
                matching_files = [
                    f
                    for f in all_txt_files_in_folder
                    if os.path.splitext(f)[0].startswith(base_filename_prefix)
                ]
                for f in sorted(matching_files):
                    if f not in processed_files_set:
                        files_to_process_in_order.append(f)
                        processed_files_set.add(f)
            unprocessed_files = [
                f for f in all_txt_files_in_folder if f not in processed_files_set
            ]
            files_to_process_in_order.extend(sorted(unprocessed_files))
        else:
            files_to_process_in_order = sorted(list(all_txt_files_in_folder))

        # --- 4. Iterar archivos y evaluar aplicaciones ---
        for txt_filename in files_to_process_in_order:
            # --- Limpiar el nombre del subdominio para agrupar (ej. quitar _1, _2) ---
            subdomain_name_raw = os.path.splitext(txt_filename)[0]
            subdomain_cleaned = re.sub(r"_\d+$", "", subdomain_name_raw)

            filepath = os.path.join(domain_path, txt_filename)

            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        parts = re.findall(r'"(.*?)"', line)

                        if len(parts) != 3:
                            continue

                        (country, bank, app_name) = parts

                        # Determinar banco y DF
                        if "BUYERBANK" in bank.upper():
                            target_df, target_choices = df_buyer, choices_buyer
                            bank_name_eval = "BuyerBank"
                        elif "BOUGHTBANK" in bank.upper():
                            target_df, target_choices = df_bought, choices_bought
                            bank_name_eval = "BoughtBank"
                        else:
                            continue

                        # Buscar el match
                        excel_match_name = find_best_match(app_name, target_choices)

                        if excel_match_name:
                            row_df = target_df[
                                target_df[APP_COLUMN_NAME_DB] == excel_match_name
                            ]
                            if row_df.empty:
                                continue
                            row = row_df.iloc[0]

                            # --- ¡EL FILTRO CLAVE! ---
                            # Solo procesar si fue marcada en main.py
                            if (
                                row is not None
                                and row.get("mostrar_en_arquitectura_target") == "Si"
                            ):
                                # Evaluar los criterios
                                resultados_evaluacion = evaluar_criterios(
                                    row, bank_name_eval, CRITERIA_DB_MAP
                                )

                                # Guardar la fila de datos
                                result_row = {
                                    "Dominio": domain_key,
                                    "Subdominio": subdomain_cleaned,  # <-- Usar el nombre limpio
                                    "Banco": bank_name_eval,
                                    "Pais": country,
                                    "Region": get_region(country),  # <-- Añadir Región
                                    "Aplicacion": app_name,
                                }
                                result_row.update(resultados_evaluacion)
                                all_app_evaluations.append(result_row)

            except Exception as e:
                print(f"Error procesando el archivo {txt_filename}: {e}")

    # --- 5. Procesar y agregar los datos ---
    if not all_app_evaluations:
        print(
            "No se encontraron aplicaciones marcadas como 'Si' para reportar. Abortando."
        )
        return

    print(
        f"Se evaluaron {len(all_app_evaluations)} aplicaciones marcadas como 'Target'."
    )
    df_raw = pd.DataFrame(all_app_evaluations)

    # Agrupar por Dominio, Subdominio, Banco y AHORA TAMBIÉN REGIÓN
    grouped = df_raw.groupby(["Dominio", "Subdominio", "Banco", "Region"])

    final_results = []

    for name_tuple, group_df in grouped:
        domain, subdomain, bank, region = name_tuple
        total_apps_in_group = len(group_df)

        result_row = {
            "Dominio": domain,
            "Subdominio": subdomain,
            "Banco": bank,
            "Region": region,  # <-- Añadir Región al resultado
            "Total Aplicaciones": total_apps_in_group,
        }

        # Calcular porcentajes para cada criterio
        for criterion in CRITERIA_LIST:
            scores_sum = 0.0
            apps_evaluadas = 0  # Apps que tienen un valor (no nulo o N/A)

            for idx, app_row in group_df.iterrows():
                compliance_text = app_row.get(criterion, "")
                score = COMPLIANCE_MAP.get(compliance_text)

                if score is not None:
                    scores_sum += score
                    apps_evaluadas += 1

            # Calcular porcentaje solo sobre apps que tenían datos
            percentage = (
                (scores_sum / apps_evaluadas) * 100 if apps_evaluadas > 0 else None
            )

            result_row[f"{criterion.capitalize()} (%)"] = percentage

        final_results.append(result_row)

    # --- 6. Aplicar Orden Personalizado ---
    df_final_report = pd.DataFrame(final_results)

    # Crear la columna de ordenamiento
    # La función get_subdomain_sort_key ya maneja los nombres limpios
    df_final_report["Subdominio_Sort_Key"] = df_final_report.apply(
        lambda row: get_subdomain_sort_key(row["Dominio"], row["Subdominio"]), axis=1
    )

    # Ordenar el reporte final (añadiendo Región)
    df_final_report = df_final_report.sort_values(
        by=["Region", "Dominio", "Subdominio_Sort_Key", "Banco"]
    )

    # Eliminar la columna de ordenamiento
    df_final_report = df_final_report.drop(columns=["Subdominio_Sort_Key"])

    # Definir el orden final de las columnas
    excel_cols = [
        "Dominio",
        "Subdominio",
        "Total Aplicaciones",
    ]  # Banco y Region se usan para la hoja
    for c in CRITERIA_LIST:
        excel_cols.append(f"{c.capitalize()} (%)")

    # --- 7. Guardar el Excel en Hojas Separadas por BANCO y REGIÓN ---

    try:
        # Asegurarse de que el directorio de salida exista
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        with pd.ExcelWriter(EXCEL_REPORT_FILE, engine="openpyxl") as writer:
            # --- BuyerBank Colombia ---
            df_report_buyer_co = df_final_report[
                (df_final_report["Banco"] == "BuyerBank")
                & (df_final_report["Region"] == "Colombia")
            ][excel_cols].reset_index(drop=True)
            df_report_buyer_co.to_excel(
                writer, sheet_name="Resumen_Buyer_CO", index=False
            )

            # --- BuyerBank CAM ---
            df_report_buyer_cam = df_final_report[
                (df_final_report["Banco"] == "BuyerBank")
                & (df_final_report["Region"] == "CAM")
            ][excel_cols].reset_index(drop=True)
            df_report_buyer_cam.to_excel(
                writer, sheet_name="Resumen_Buyer_CAM", index=False
            )

            # --- BoughtBank Colombia ---
            df_report_bought_co = df_final_report[
                (df_final_report["Banco"] == "BoughtBank")
                & (df_final_report["Region"] == "Colombia")
            ][excel_cols].reset_index(drop=True)
            df_report_bought_co.to_excel(
                writer, sheet_name="Resumen_Bought_CO", index=False
            )

            # --- BoughtBank CAM ---
            df_report_bought_cam = df_final_report[
                (df_final_report["Banco"] == "BoughtBank")
                & (df_final_report["Region"] == "CAM")
            ][excel_cols].reset_index(drop=True)
            df_report_bought_cam.to_excel(
                writer, sheet_name="Resumen_Bought_CAM", index=False
            )

            # --- Hojas de Datos Crudos ---
            df_raw[
                (df_raw["Banco"] == "BuyerBank") & (df_raw["Region"] == "Colombia")
            ].to_excel(writer, sheet_name="Datos_Raw_Buyer_CO", index=False)

            df_raw[
                (df_raw["Banco"] == "BuyerBank") & (df_raw["Region"] == "CAM")
            ].to_excel(writer, sheet_name="Datos_Raw_Buyer_CAM", index=False)

            df_raw[
                (df_raw["Banco"] == "BoughtBank") & (df_raw["Region"] == "Colombia")
            ].to_excel(writer, sheet_name="Datos_Raw_Bought_CO", index=False)

            df_raw[
                (df_raw["Banco"] == "BoughtBank") & (df_raw["Region"] == "CAM")
            ].to_excel(writer, sheet_name="Datos_Raw_Bought_CAM", index=False)

        print(f"\n✅ ¡Éxito! Reporte de cumplimiento guardado en:")
        print(f"{EXCEL_REPORT_FILE}")

    except Exception as e:
        print(f"\n❌ ERROR al guardar el archivo Excel: {e}")
        print(
            "Asegúrate de no tener el archivo 'Reporte_Cumplimiento_Target.xlsx' abierto."
        )


if __name__ == "__main__":
    generar_reporte()
