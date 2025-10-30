"""
Script para encontrar sugerencias de 'LIKE' para las aplicaciones
que no se encontraron (pendientes.txt) durante la ejecución de main.py.
"""

import sys
import os
import re

try:
    from sqlalchemy import create_engine, text
    from sqlalchemy.exc import OperationalError
except ImportError:
    print("Error: 'sqlalchemy' o 'mysql-connector-python' no encontrados.")
    print("Por favor, instálalos: pip install sqlalchemy mysql-connector-python")
    sys.exit(1)

# --- CONFIGURACIÓN DE LA BASE DE DATOS ---
# (Debe coincidir con tu docker-compose.yml y db_loader.py)
DB_USER = "root"
DB_PASS = "admin"
DB_HOST = "localhost"
DB_PORT = "3306"
DB_NAME = "bank_app_analysis"

DB_URL = f"mysql+mysqlconnector://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}"

# Nombres de tablas y columnas
TABLE_BUYER_BANK = "aplicaciones_buyer_bank"
TABLE_BOUGHT_BANK = "aplicaciones_bought_bank"
APP_COLUMN_NAME_DB = "aplicacion_sistema"

# --- ARCHIVOS DE ENTRADA Y SALIDA ---
# Usar el nuevo subdirectorio
PENDING_DIR = os.path.join("outputs", "pendientes")
PENDIENTES_FILE = os.path.join(PENDING_DIR, "pendientes.txt")
SUGERENCIAS_FILE = os.path.join(PENDING_DIR, "sugerencias_pendientes.txt")


def parse_pendientes():
    """
    Lee el archivo pendientes.txt y extrae los nombres de las aplicaciones
    para cada banco, evitando duplicados.
    """
    pending_buyer = set()
    pending_bought = set()

    try:
        with open(PENDIENTES_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("---"):
                    continue

                parts = re.findall(r'"(.*?)"', line)

                if len(parts) == 3:
                    _country, bank, app_name = parts
                    app_name = app_name.strip()
                    if "BuyerBank" in bank:
                        pending_buyer.add(app_name)
                    elif "BoughtBank" in bank:
                        pending_bought.add(app_name)

    except FileNotFoundError:
        print(f"Error: No se encontró el archivo '{PENDIENTES_FILE}'.")
        print(
            "Asegúrate de ejecutar main.py primero para generar el archivo de pendientes."
        )
        return None, None
    except Exception as e:
        print(f"Error leyendo '{PENDIENTES_FILE}': {e}")
        return None, None

    return pending_buyer, pending_bought


def find_suggestions(engine, table_name, app_names_set, writer):
    """
    Toma un conjunto de nombres de aplicaciones pendientes y ejecuta una
    consulta LIKE contra la tabla especificada.
    """

    # Ordenamos la lista para un reporte más limpio
    for app_name in sorted(list(app_names_set)):
        # Preparamos el patrón de búsqueda LIKE
        like_pattern = f"%{app_name}%"

        query = text(f"""
            SELECT {APP_COLUMN_NAME_DB} 
            FROM {table_name} 
            WHERE {APP_COLUMN_NAME_DB} LIKE :pattern
            LIMIT 10
        """)

        try:
            with engine.connect() as conn:
                results = conn.execute(query, {"pattern": like_pattern}).fetchall()

            if results:
                # Si encontramos coincidencias, las escribimos en el archivo
                writer.write(f'\n--- PENDIENTE: "{app_name}" ---\n')
                for row in results:
                    writer.write(f'  -> POSIBLE MATCH: "{row[0]}"\n')

        except Exception as e:
            print(f"Error al buscar sugerencias para '{app_name}': {e}")


def main():
    print(f"Leyendo '{PENDIENTES_FILE}'...")
    pending_buyer, pending_bought = parse_pendientes()

    if pending_buyer is None:
        return  # Error durante el parseo

    print(
        f"Encontradas {len(pending_buyer)} apps pendientes (Buyer) y {len(pending_bought)} (Bought)."
    )

    try:
        engine = create_engine(DB_URL)
        # Probar conexión
        with engine.connect() as conn:
            pass
        print("Conexión a la base de datos exitosa.")
    except Exception as e:
        print(f"Error: No se pudo conectar a la base de datos.")
        print(f"Detalle: {e}")
        return

    # Asegurarse de que el directorio de salida exista
    os.makedirs(PENDING_DIR, exist_ok=True)

    # Escribir el nuevo archivo de sugerencias
    try:
        with open(SUGERENCIAS_FILE, "w", encoding="utf-8") as writer:
            print(f"Buscando sugerencias para {TABLE_BUYER_BANK}...")
            writer.write("=" * 40 + "\n")
            writer.write(f" SUGERENCIAS PARA: {TABLE_BUYER_BANK}\n")
            writer.write("=" * 40 + "\n")
            find_suggestions(engine, TABLE_BUYER_BANK, pending_buyer, writer)

            print(f"Buscando sugerencias para {TABLE_BOUGHT_BANK}...")
            writer.write("\n\n")
            writer.write("=" * 40 + "\n")
            writer.write(f" SUGERENCIAS PARA: {TABLE_BOUGHT_BANK}\n")
            writer.write("=" * 40 + "\n")
            find_suggestions(engine, TABLE_BOUGHT_BANK, pending_bought, writer)

        print(
            f"\n¡Éxito! Se ha generado el archivo '{SUGERENCIAS_FILE}' con las posibles coincidencias."
        )

    except Exception as e:
        print(f"Error al escribir el archivo de sugerencias: {e}")


if __name__ == "__main__":
    main()
