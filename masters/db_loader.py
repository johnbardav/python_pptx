"""
Handles loading data from the MySQL database
for the Bank App Analysis project.
"""

import sys
import pandas as pd

try:
    from sqlalchemy import create_engine
    from sqlalchemy.exc import OperationalError
except ImportError:
    print("Error: 'sqlalchemy' or 'mysql-connector-python' not found.")
    print("Please install them: pip install sqlalchemy mysql-connector-python")
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

# Table names defined in load_database.py
TABLE_BUYER_BANK = "aplicaciones_buyer_bank"
TABLE_BOUGHT_BANK = "aplicaciones_bought_bank"


def load_data_from_db():
    """
    Connects to the MySQL database and loads the application tables
    into pandas DataFrames.

    Returns:
        A tuple containing: (df_buyer, df_bought, engine)
        Returns (None, None, None) if the connection fails.
    """
    try:
        engine = create_engine(DB_URL)
        with engine.connect() as conn:
            pass  # Test connection
        print("Database connection successful.")
    except OperationalError as e:
        print(f"\nError: Could not connect to the database at {DB_HOST}:{DB_PORT}.")
        print("Please ensure the Docker container is running ('docker-compose up -d').")
        print(f"Details: {e}")
        return None, None, None
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
        return None, None, None

    try:
        print(f"Reading table '{TABLE_BUYER_BANK}'...")
        df_buyer = pd.read_sql(f"SELECT * FROM {TABLE_BUYER_BANK}", engine)

        print(f"Reading table '{TABLE_BOUGHT_BANK}'...")
        df_bought = pd.read_sql(f"SELECT * FROM {TABLE_BOUGHT_BANK}", engine)

        return df_buyer, df_bought, engine

    except Exception as e:
        print(f"Error reading data from tables: {e}")
        return None, None, None
