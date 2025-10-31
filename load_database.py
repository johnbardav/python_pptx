"""
Script principal para cargar datos desde archivos Excel maestros
y el archivo 'Consolidado.xlsx' a la base de datos MySQL (gestionada por Docker).
"""

import sys
import time
import re
import pandas as pd

try:
    from sqlalchemy import create_engine, text
    from sqlalchemy.exc import OperationalError
except ImportError:
    print("Error: 'sqlalchemy' or 'mysql-connector-python' not found.")
    print("Please install them: pip install sqlalchemy mysql-connector-python")
    sys.exit(1)

# Importar el cargador de Excel existente
try:
    from masters import load_master_excels
except ImportError as e:
    print(f"Error importing 'masters' module: {e}")
    print("Ensure you are running this script from the project root directory.")
    sys.exit(1)


# --- Database Configuration ---
# These must match your docker-compose.yml
DB_USER = "root"
DB_PASS = "admin"
DB_HOST = "localhost"
DB_PORT = "3306"
DB_NAME = "bank_app_analysis"

# Connection String
DB_URL = f"mysql+mysqlconnector://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}"

# Define table names
TABLE_BUYER_BANK = "aplicaciones_buyer_bank"
TABLE_BOUGHT_BANK = "aplicaciones_bought_bank"
TABLE_CONSOLIDADO = "aplicaciones_consolidadas"

# Define file paths
CONSOLIDADO_FILE_PATH = "inputs/Consolidado.xlsx"

# Límite de MySQL (64) menos espacio para sufijos (ej. "_10")
MAX_COL_LENGTH = 61


def wait_for_db(engine, retries=15, wait_time=5):
    """
    Waits for the Docker database container to be ready
    before attempting to connect.
    """
    print("Attempting to connect to the database...")
    for i in range(retries):
        try:
            with engine.connect() as conn:
                conn.execute(text("SELECT 1"))
            print("Database connection successful!")
            return True
        except OperationalError:
            print(f"Database not ready... retrying in {wait_time}s ({i + 1}/{retries})")
            time.sleep(wait_time)

    print(f"Error: Could not connect to the database after {retries} retries.")
    return False


def deduplicate_columns(df):
    """
    Renames duplicate columns by appending a suffix (_1, _2, etc.).
    e.g., ['col', 'col'] becomes ['col', 'col_1']
    """
    new_cols = []
    counts = {}
    for col in df.columns:
        counts[col] = counts.get(col, 0)  # Get current count, default to 0
        if counts[col] > 0:  # If this is a duplicate
            new_name = f"{col}_{counts[col]}"  # Append the count
        else:
            new_name = col  # First time, use original name

        new_cols.append(new_name)
        counts[col] += 1  # Increment the count for this name

    df.columns = new_cols
    return df


def clean_and_truncate_cols(cols):
    """
    Cleans and truncates a list of column names for MySQL compatibility.
    """
    cleaned_cols = []
    for col in cols:
        # 1. Clean spaces/hyphens/special chars and convert to lowercase
        new_col = (
            str(col)
            .lower()  # <-- FIX: Convert to lowercase
            .replace(" ", "_")
            .replace("-", "_")
            .replace("/", "_")
            .replace("(", "")
            .replace(")", "")
            .replace("¿", "")
            .replace("?", "")
        )
        # 2. Truncate if necessary (leaving space for suffix)
        if len(new_col) > MAX_COL_LENGTH:
            new_col = new_col[:MAX_COL_LENGTH]
        cleaned_cols.append(new_col)
    return cleaned_cols


def load_data_to_db():
    """
    Main function to orchestrate the ETL process:
    1. Load original Excel files using 'load_master_excels'.
    2. Clean and write DataFrames to MySQL tables (buyer_bank, bought_bank).
    3. Load 'Consolidado.xlsx'.
    4. Clean and write DataFrame to MySQL table (aplicaciones_consolidadas).
    """

    engine = create_engine(DB_URL)

    # 1. Wait for Docker container to be online
    if not wait_for_db(engine):
        print("Aborting database load.")
        return

    # --- PART 1: Load Original Master Excels ---
    print("Loading master Excel files (Buyer/Bought) into memory...")
    (df_buyer, choices_buyer, df_bought, choices_bought) = load_master_excels()

    if df_buyer is None or df_bought is None:
        print("Could not load one or more master Excel files. Aborting.")
        return

    print("Master Excel files loaded successfully.")

    # 3. Data Cleaning Pipeline (Buyer/Bought)
    print("Cleaning Buyer/Bought data...")
    # STEP 3.1: Remove blank columns
    df_buyer = df_buyer.loc[:, df_buyer.columns != ""]
    df_bought = df_bought.loc[:, df_bought.columns != ""]

    # STEP 3.2: Clean and Truncate names
    df_buyer.columns = clean_and_truncate_cols(df_buyer.columns)
    df_bought.columns = clean_and_truncate_cols(df_bought.columns)
    print(f"Cleaned and truncated column names to {MAX_COL_LENGTH} chars.")

    # STEP 3.3: De-duplicate column names
    df_buyer = deduplicate_columns(df_buyer)
    df_bought = deduplicate_columns(df_bought)
    print("Renamed any duplicate columns (post-truncation).")

    # 4. Load DataFrames into MySQL tables (Buyer/Bought)
    try:
        print(f"Writing Buyer Bank data to table '{TABLE_BUYER_BANK}'...")
        df_buyer.to_sql(TABLE_BUYER_BANK, engine, if_exists="replace", index=False)
        print("...Buyer Bank data written successfully.")

        print(f"Writing Bought Bank data to table '{TABLE_BOUGHT_BANK}'...")
        df_bought.to_sql(TABLE_BOUGHT_BANK, engine, if_exists="replace", index=False)
        print("...Bought Bank data written successfully.")

    except Exception as e:
        print(f"Error during .to_sql() operation for Buyer/Bought: {e}")
        print("Please check database permissions and data types.")
        return

    print("\n--- Master tables load complete! ---")

    # --- PART 2: Load Consolidado Excel ---
    print(f"\nLoading Excel file '{CONSOLIDADO_FILE_PATH}' into memory...")
    try:
        # Los encabezados están en la fila 1 (índice 0)
        # Los datos comienzan en la fila 18 (índice 17)
        # Omitimos las filas de descripción intermedias (índices 1 a 16).
        df_consolidado = pd.read_excel(
            CONSOLIDADO_FILE_PATH, header=0, skiprows=range(1, 17)
        )

    except FileNotFoundError:
        print(f"Error: File not found at '{CONSOLIDADO_FILE_PATH}'.")
        print("Ensure the file 'Consolidado.xlsx' is in the 'inputs/' directory.")
        return
    except Exception as e:
        print(f"Error loading Excel file '{CONSOLIDADO_FILE_PATH}': {e}")
        return

    print("Consolidado Excel file loaded successfully.")

    # Data Cleaning Pipeline (Consolidado)
    print("Cleaning Consolidado data...")
    # STEP A: Remove blank columns
    original_cols_consol = len(df_consolidado.columns)
    df_consolidado = df_consolidado.loc[:, df_consolidado.columns != ""]
    print(
        f"Removed {original_cols_consol - len(df_consolidado.columns)} blank columns from Consolidado."
    )

    # STEP B: Clean and Truncate names
    df_consolidado.columns = clean_and_truncate_cols(df_consolidado.columns)
    print(f"Cleaned and truncated column names to {MAX_COL_LENGTH} chars.")

    # STEP C: De-duplicate column names
    df_consolidado = deduplicate_columns(df_consolidado)
    print("Renamed any duplicate columns (post-truncation).")

    print(f"Final Consolidado columns: {list(df_consolidado.columns)}")

    # Load DataFrame into MySQL table (Consolidado)
    try:
        print(f"Writing Consolidado data to table '{TABLE_CONSOLIDADO}'...")
        df_consolidado.to_sql(
            TABLE_CONSOLIDADO, engine, if_exists="replace", index=False
        )
        print("...Consolidado data written successfully.")

    except Exception as e:
        print(f"Error during .to_sql() operation for Consolidado: {e}")
        print("Please check database permissions and data types.")
        return

    print("\nDatabase load complete! All files are now in MySQL.")


# --- Main execution block ---
if __name__ == "__main__":
    load_data_to_db()
