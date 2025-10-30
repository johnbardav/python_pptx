"""
Handles the loading and pre-processing of the master Excel files
for the Bank App Analysis project.
"""

import pandas as pd
import re
import os
import sys

try:
    from unidecode import unidecode
except ImportError:
    print("Error: The 'unidecode' library is not installed.")
    print("Please install it using: pip install unidecode")
    sys.exit(1)

# --- Constants based on reference script ---
INPUT_FOLDER = "inputs"
APP_COLUMN_NAME = "aplicacion sistema"

# --- Archivos y Hojas ---
# Nombres de archivo genéricos que DEBES usar en tu carpeta 'inputs'
FILE_BUYER_BANK_NAME = "master_buyer_bank.xlsx"
SHEET_BUYER_BANK = "Applications"

FILE_BOUGHT_BANK_NAME = "master_bought_bank.xlsx"
SHEET_BOUGHT_BANK = "Applications"


def normalize_string(text: str) -> str:
    """
    Cleans and normalizes a string (e.g., app names, columns).
    Based on the logic from 1_generador_madurez_y_reportes.py.
    """
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


def load_master_excels():
    """
    Loads and processes the BuyerBank and BoughtBank master Excel files
    from the 'inputs' directory.

    Returns:
        A tuple containing: (df_buyer, choices_buyer, df_bought, choices_bought)
        Returns (None, None, None, None) if any file is not found.
    """
    df_buyer, choices_buyer = None, {}
    df_bought, choices_bought = None, {}

    file_buyer_path = os.path.join(INPUT_FOLDER, FILE_BUYER_BANK_NAME)
    file_bought_path = os.path.join(INPUT_FOLDER, FILE_BOUGHT_BANK_NAME)

    # --- Load Buyer Bank File ---
    try:
        print(f"Loading Buyer Bank file: {FILE_BUYER_BANK_NAME}...")
        df_buyer = pd.read_excel(file_buyer_path, sheet_name=SHEET_BUYER_BANK)
        df_buyer["banco"] = "BuyerBank"
        df_buyer.columns = [normalize_string(col) for col in df_buyer.columns]

        if APP_COLUMN_NAME not in df_buyer.columns:
            print(
                f"Error: Required column '{APP_COLUMN_NAME}' not found in Buyer Bank file after normalization."
            )
            return None, None, None, None

        choices_buyer = {
            normalize_string(name): name
            for name in df_buyer[APP_COLUMN_NAME].dropna().unique()
        }
        print("Buyer Bank file loaded successfully.")

    except FileNotFoundError:
        print(f"Error: Buyer Bank file not found at {file_buyer_path}")
        return None, None, None, None
    except Exception as e:
        print(f"Error loading Buyer Bank file: {e}")
        return None, None, None, None

    # --- Load Bought Bank File ---
    try:
        print(f"Loading Bought Bank file: {FILE_BOUGHT_BANK_NAME}...")
        df_bought = pd.read_excel(file_bought_path, sheet_name=SHEET_BOUGHT_BANK)
        df_bought["banco"] = "BoughtBank"
        df_bought.columns = [normalize_string(col) for col in df_bought.columns]

        if APP_COLUMN_NAME not in df_bought.columns:
            print(
                f"Error: Required column '{APP_COLUMN_NAME}' not found in Bought Bank file after normalization."
            )
            return None, None, None, None

        choices_bought = {
            normalize_string(name): name
            for name in df_bought[APP_COLUMN_NAME].dropna().unique()
        }
        print("Bought Bank file loaded successfully.")

    except FileNotFoundError:
        print(f"Error: Bought Bank file not found at {file_bought_path}")
        return None, None, None, None
    except Exception as e:
        print(f"Error loading Bought Bank file: {e}")
        return None, None, None, None

    return df_buyer, choices_buyer, df_bought, choices_bought
