"""
Main script to generate the PowerPoint presentation
for the Bank App Analysis project.
"""

import sys
import os

try:
    # PPTX/Util
    from pptx import Presentation
    from pptx.util import Cm

    # Funciones de nuestros 'masters'
    from masters import load_data_from_db, generate_slide_for_txt, normalize_string
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
    # (Lógica adaptada de 1_generador_madurez_y_reportes.py)
    print("Creating fuzzy-matching lookups...")
    choices_buyer = {
        normalize_string(name): name
        for name in df_buyer[APP_COLUMN_NAME_DB].dropna().unique()
    }
    choices_bought = {
        normalize_string(name): name
        for name in df_bought[APP_COLUMN_NAME_DB].dropna().unique()
    }

    # --- 3. Find all .txt files in 'inputs' ---
    files_to_process = []
    for dirpath, _, filenames in os.walk(INPUT_FOLDER):
        for filename in filenames:
            if filename.endswith(".txt"):
                files_to_process.append(os.path.join(dirpath, filename))

    if not files_to_process:
        print(
            f"🟡 Aviso: No se encontraron archivos .txt en '{INPUT_FOLDER}' y sus subdirectorios."
        )
        print("No slides will be generated.")
        return

    print(f"\nFound {len(files_to_process)} .txt files to process...")

    # --- 4. Process each .txt file to generate a .pptx ---
    for filepath in files_to_process:
        try:
            generate_slide_for_txt(
                filepath,
                df_buyer,
                choices_buyer,
                df_bought,
                choices_bought,
            )
        except Exception as e:
            print(f"❌ ERROR Fatal procesando el archivo '{filepath}': {e}")

    print("\nProcess finished.")


# --- Main execution block ---
if __name__ == "__main__":
    # Create the output directory if it doesn't exist
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        print(f"Created directory: {OUTPUT_FOLDER}")

    main_orchestrator()
