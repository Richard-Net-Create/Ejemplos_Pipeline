"""
PIPELINE DE LIMPIEZA SQL + CARGA A POSTGRESQL
=============================================
Lee un Excel/CSV, lo limpia con una consulta SQL (SQLite en memoria)
y carga el resultado en una tabla de PostgreSQL.

Configuración por variables de entorno (ver .env.example):
    INPUT_PATH, SHEET_NAME, PG_USER, PG_PASSWORD, PG_HOST, PG_PORT,
    PG_DATABASE, PG_TABLE, IF_EXISTS, OUTPUT_LOCAL

Uso:
    python pipeline_postgres.py
"""

import os
import sqlite3
import sys
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL

INPUT_PATH = os.environ.get("INPUT_PATH", "datos/Nomina_Matricula_ejemplo.xlsx")
SHEET_NAME = os.environ.get("SHEET_NAME", "Nomina Matricula")
OUTPUT_LOCAL = os.environ.get("OUTPUT_LOCAL")  # opcional: copia local .xlsx/.csv

PG_USER = os.environ.get("PG_USER", "postgres")
PG_PASSWORD = os.environ.get("PG_PASSWORD")
PG_HOST = os.environ.get("PG_HOST", "localhost")
PG_PORT = int(os.environ.get("PG_PORT", "5432"))
PG_DATABASE = os.environ.get("PG_DATABASE", "colegio")
PG_TABLE = os.environ.get("PG_TABLE", "nomina_matricula")
IF_EXISTS = os.environ.get("IF_EXISTS", "append")  # append | replace | fail

CLEAN_SQL = """
    SELECT DISTINCT
        "N° Orden",
        "N° Mat.",
        TRIM("Apellidos y Nombres") AS "Apellidos y Nombres",
        "Edad",
        "Sexo",
        "Repite",
        "Cond.",
        COALESCE(NULLIF(TRIM("Curso de Cargo"), ''), 'Sin asignar') AS "Curso de Cargo",
        "Centro Educativo de Procedencia"
    FROM raw
    WHERE "Apellidos y Nombres" IS NOT NULL
      AND TRIM("Apellidos y Nombres") <> ''
"""

COLUMN_MAP = {
    "N° Orden": "orden",
    "N° Mat.": "num_matricula",
    "Apellidos y Nombres": "apellidos_nombres",
    "Edad": "edad",
    "Sexo": "sexo",
    "Repite": "repite",
    "Cond.": "condicion",
    "Curso de Cargo": "curso_cargo",
    "Centro Educativo de Procedencia": "centro_educativo",
}


def load_source(path: str, sheet_name) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path)
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path, sheet_name=sheet_name)
    sys.exit(f"Formato no soportado: {ext}")


def clean_with_sql(df: pd.DataFrame) -> pd.DataFrame:
    conn = sqlite3.connect(":memory:")
    df.to_sql("raw", conn, index=False, if_exists="replace")
    limpio = pd.read_sql_query(CLEAN_SQL, conn)
    conn.close()
    return limpio


def upload_to_postgres(df: pd.DataFrame) -> None:
    if not PG_PASSWORD:
        sys.exit("Define la variable de entorno PG_PASSWORD (ver .env.example).")
    # URL.create escapa bien caracteres especiales de la contraseña
    url = URL.create(
        "postgresql+pg8000",
        username=PG_USER,
        password=PG_PASSWORD,
        host=PG_HOST,
        port=PG_PORT,
        database=PG_DATABASE,
    )
    engine = create_engine(url)
    df.to_sql(PG_TABLE, engine, index=False, if_exists=IF_EXISTS)
    engine.dispose()


def main():
    print(f"1) Leyendo origen: {INPUT_PATH}")
    df_in = load_source(INPUT_PATH, SHEET_NAME)
    print(f"   -> {len(df_in)} filas cargadas")

    print("2) Limpiando con SQL...")
    df_clean = clean_with_sql(df_in).rename(columns=COLUMN_MAP)
    print(f"   -> {len(df_clean)} filas tras limpiar "
          f"({len(df_in) - len(df_clean)} descartadas)")

    print(f"3) Subiendo a PostgreSQL: {PG_DATABASE}.{PG_TABLE} (if_exists={IF_EXISTS})")
    upload_to_postgres(df_clean)
    print("   -> Carga completa")

    if OUTPUT_LOCAL:
        if OUTPUT_LOCAL.lower().endswith(".csv"):
            df_clean.to_csv(OUTPUT_LOCAL, index=False)
        else:
            df_clean.to_excel(OUTPUT_LOCAL, index=False)
        print(f"4) Copia local guardada en: {OUTPUT_LOCAL}")

    print("Listo.")


if __name__ == "__main__":
    main()
