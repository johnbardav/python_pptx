"""
Main script to load data from master Excel files into
the MySQL database (managed by Docker).
"""

import sys
import time
import re

try:
    from sqlalchemy import create_engine, text
    from sqlalchemy.exc import OperationalError
except ImportError:
    print("Error: 'sqlalchemy' or 'mysql-connector-python' not found.")
    print("Please install them: pip install sqlalchemy mysql-connector-python")
    sys.exit(1)

# Import our existing Excel loader
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


def load_data_to_db():
    """
    Main function to orchestrate the ETL process:
    1. Load Excel files using 'load_master_excels'.
    2. Clean column names as requested (spaces to underscores).
    3. Write DataFrames to MySQL tables.
    """

    engine = create_engine(DB_URL)

    # 1. Wait for Docker container to be online
    if not wait_for_db(engine):
        print("Aborting database load.")
        return

    # 2. Load data from Excel using our master function
    print("Loading master Excel files into memory...")
    (df_buyer, choices_buyer, df_bought, choices_bought) = load_master_excels()

    if df_buyer is None or df_bought is None:
        print("Could not load one or more Excel files. Aborting.")
        return

    print("Excel files loaded successfully.")

    # 3. Clean column names
    df_buyer.columns = [
        col.replace(" ", "_").replace("-", "_") for col in df_buyer.columns
    ]
    df_bought.columns = [
        col.replace(" ", "_").replace("-", "_") for col in df_bought.columns
    ]

    print(f"Cleaned Buyer Bank columns: {list(df_buyer.columns)}")
    print(f"Cleaned Bought Bank columns: {list(df_bought.columns)}")

    # 4. Load DataFrames into MySQL tables
    try:
        print(f"Writing Buyer Bank data to table '{TABLE_BUYER_BANK}'...")
        df_buyer.to_sql(TABLE_BUYER_BANK, engine, if_exists="replace", index=False)
        print("...Buyer Bank data written successfully.")

        print(f"Writing Bought Bank data to table '{TABLE_BOUGHT_BANK}'...")
        df_bought.to_sql(TABLE_BOUGHT_BANK, engine, if_exists="replace", index=False)
        print("...Bought Bank data written successfully.")

        print("\nDatabase load complete! Data is now in MySQL.")

    except Exception as e:
        print(f"Error during .to_sql() operation: {e}")
        print("Please check database permissions and data types.")


# --- Main execution block ---
if __name__ == "__main__":
    load_data_to_db()
