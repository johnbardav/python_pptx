"""
Main script to generate the PowerPoint presentation
for the Bank App Analysis project.
"""

import sys
import os
import re
from collections import defaultdict

try:
    # PPTX/Util
    from pptx import Presentation
    from pptx.util import Cm

    # Funciones de nuestros 'masters'
    from masters import (
        load_data_from_db,
        generate_slide_for_subdomain,
        normalize_string,
    )

    # Importar configuración
    from config import CUSTOM_SORT_ORDER, CRITERIA_DB_MAP

except ImportError as e:
    print(f"Import Error: {e}")
    print("Please ensure all dependencies from 'requirements.txt' are installed.")
    print("Run 'install.bat' or 'install.sh'.")
    sys.exit(1)
except Exception as e:
    print(f"An unexpected error occurred during imports: {e}")
    sys.exit(1)

# --- Constantes ---
INPUT_FOLDER = "inputs"
OUTPUT_FOLDER = "outputs"
# El nombre de la columna de aplicación en la BD (ya limpio)
APP_COLUMN_NAME_DB = "aplicacion_sistema"


def main_orchestrator():
    """
    Orquesta todo el proceso de generación de slides.
    """

    # --- 1. Load Data from Database ---
    print("\nConnecting to database and loading data...")
    (
        df_buyer,
        df_bought,
    ) = load_data_from_db()

    if df_buyer is None or df_bought is None:
        print("Could not load data from database. Aborting presentation build.")
        return

    print(
        f"Loaded {len(df_buyer)} Buyer Bank apps and {len(df_bought)} Bought Bank apps from DB."
    )

    # --- 2. Create 'choices' lookup dictionaries for fuzzy matching ---
    print("Creating fuzzy-matching lookups...")
    choices_buyer = {
        normalize_string(name): name
        for name in df_buyer[APP_COLUMN_NAME_DB].dropna().unique()
    }
    choices_bought = {
        normalize_string(name): name
        for name in df_bought[APP_COLUMN_NAME_DB].dropna().unique()
    }

    # --- 3. Find domains (subfolders) in INPUT_FOLDER ---
    print(f"Scanning '{INPUT_FOLDER}' for domain folders...")
    domain_folders = [
        f
        for f in os.listdir(INPUT_FOLDER)
        if os.path.isdir(os.path.join(INPUT_FOLDER, f))
    ]

    if not domain_folders:
        print(
            f"🟡 Aviso: No se encontraron subcarpetas de dominio en '{INPUT_FOLDER}'."
        )
        print(
            "Asegúrate de que tu estructura sea 'inputs/Canales/', 'inputs/Soporte/', etc."
        )
        return

    print(f"\nFound {len(domain_folders)} domains to process...")

    # --- 4. Process each DOMAIN to generate one PPTX file ---
    for domain_name in domain_folders:
        print(f"\nProcessing Domain: {domain_name}...")

        # 4.1. Create a new presentation for this domain
        prs = Presentation()
        prs.slide_width = Cm(33.87)
        prs.slide_height = Cm(19.05)

        domain_path = os.path.join(INPUT_FOLDER, domain_name)
        domain_key = domain_name.lower()  # ej: "canales"

        # Obtener todos los .txt de la carpeta (para referencia)
        all_txt_files_in_folder = {
            f for f in os.listdir(domain_path) if f.endswith(".txt")
        }

        # --- INICIO DE LA MODIFICACIÓN: Lógica de Ordenamiento ---

        files_to_process_in_order = []
        processed_files_set = set()

        # Obtener la lista de ordenamiento desde config.py
        ordered_base_names = CUSTOM_SORT_ORDER.get(domain_key)

        if ordered_base_names:
            print(f"  Aplicando orden personalizado para '{domain_name}'...")

            # 1. Iterar sobre el ORDEN PERSONALIZADO
            for base_name in ordered_base_names:
                # Construir el nombre de archivo base, ej: "canales_sitio_publico"
                base_filename_prefix = f"{domain_key}_{base_name}"

                # Encontrar todos los archivos que coinciden (para _1.txt, _2.txt, etc.)
                matching_files = [
                    f
                    for f in all_txt_files_in_folder
                    if os.path.splitext(f)[0].startswith(base_filename_prefix)
                ]

                # Ordenar los coincidentes (ej. _1.txt antes de _2.txt) y agregarlos
                for f in sorted(matching_files):
                    if f not in processed_files_set:
                        files_to_process_in_order.append(f)
                        processed_files_set.add(f)

            # 2. Agregar archivos que se encontraron pero no estaban en la lista de orden
            unprocessed_files = [
                f for f in all_txt_files_in_folder if f not in processed_files_set
            ]
            if unprocessed_files:
                print(
                    f"  Aviso: Se encontraron {len(unprocessed_files)} subdominios no definidos en el orden personalizado. Añadiendo al final."
                )
                files_to_process_in_order.extend(sorted(unprocessed_files))

        else:
            # Fallback: Si el dominio no está en el diccionario, ordenar alfabéticamente
            print(
                f"  Aviso: No se encontró orden personalizado para '{domain_name}'. Usando orden alfabético."
            )
            files_to_process_in_order = sorted(list(all_txt_files_in_folder))

        # --- FIN DE LA LÓGICA DE ORDENAMIENTO ---

        if not files_to_process_in_order:
            print(
                f"  🟡 Aviso: No se encontraron archivos .txt en '{domain_path}'. Saltando."
            )
            continue

        print(
            f"  Generando {len(files_to_process_in_order)} slides en el orden definido..."
        )

        # 4.2. Loop through each subdomain .txt file and create ONE SLIDE
        for txt_filename in files_to_process_in_order:
            # Limpiar el nombre del archivo para usarlo como título del slide
            subdomain_name_raw = os.path.splitext(txt_filename)[0]
            subdomain_title = (
                subdomain_name_raw.replace("_", " ").replace("-", " ").title()
            )
            filepath = os.path.join(domain_path, txt_filename)

            app_lines_data = []
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        parts = re.findall(r'"(.*?)"', line)
                        if len(parts) == 3:
                            app_lines_data.append(tuple(parts))
                        else:
                            print(
                                f"      ⚠️  Línea ignorada (formato): '{line}' en {txt_filename}"
                            )

                # 4.3. Call the generator to ADD ONE SLIDE
                generate_slide_for_subdomain(
                    prs,
                    subdomain_title,  # Title for the slide
                    app_lines_data,
                    df_buyer,
                    choices_buyer,
                    df_bought,
                    choices_bought,
                    CRITERIA_DB_MAP,  # <-- Pasamos el mapa de configuración
                )
            except Exception as e:
                print(f"❌ ERROR Fatal procesando el subdominio '{txt_filename}': {e}")

        # 4.4. Save the completed PPTX for this domain
        output_filename = os.path.join(OUTPUT_FOLDER, f"{domain_name}.pptx")
        try:
            prs.save(output_filename)
            print(
                f"\n  ✅ ¡Éxito! Presentación de dominio guardada como '{output_filename}'!"
            )
        except PermissionError:
            print(
                f"\n  Error: No se pudo guardar '{output_filename}'. Archivo abierto?"
            )
        except Exception as e:
            print(f"\n  An unexpected error occurred saving '{output_filename}': {e}")

    print("\nProcess finished.")


# --- Main execution block ---
if __name__ == "__main__":
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        print(f"Created directory: {OUTPUT_FOLDER}")

    main_orchestrator()
