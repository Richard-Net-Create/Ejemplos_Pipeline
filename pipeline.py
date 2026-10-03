"""
PIPELINE DE LIMPIEZA DE DATOS (CSV/Excel -> SQL -> CSV/Excel)
=============================================================
Carga un archivo como tabla en SQLite (en memoria), aplica una consulta
SQL de limpieza (duplicados, nulos, condiciones) y exporta el resultado.

Uso:
    python pipeline.py --entrada datos/archivo.xlsx --salida salida/limpio.xlsx
    python pipeline.py --entrada datos/archivo.csv  --salida salida/limpio.csv
"""

import argparse
import sqlite3
import sys
from pathlib import Path

import pandas as pd

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


def load_source(path: str, sheet_name) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path)
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path, sheet_name=sheet_name)
    sys.exit(f"Formato no soportado: {ext} (usa .csv, .xlsx o .xls)")


def save_output(df: pd.DataFrame, path: str) -> None:
    ext = Path(path).suffix.lower()
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    if ext == ".csv":
        df.to_csv(path, index=False)
    elif ext in (".xlsx", ".xls"):
        df.to_excel(path, index=False)
    else:
        sys.exit(f"Formato de salida no soportado: {ext}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--entrada", required=True, help="Archivo .csv o .xlsx")
    parser.add_argument("--salida", required=True, help="Archivo de salida .csv o .xlsx")
    parser.add_argument("--hoja", default="Nomina Matricula", help="Hoja (solo Excel)")
    args = parser.parse_args()

    print(f"1) Leyendo origen: {args.entrada}")
    df_in = load_source(args.entrada, args.hoja)
    n_in = len(df_in)
    print(f"   -> {n_in} filas cargadas")

    print("2) Ejecutando limpieza SQL (SQLite en memoria)...")
    conn = sqlite3.connect(":memory:")
    df_in.to_sql("raw", conn, index=False, if_exists="replace")
    df_out = pd.read_sql_query(CLEAN_SQL, conn)
    conn.close()
    print(f"   -> {len(df_out)} filas tras limpiar ({n_in - len(df_out)} descartadas)")

    print(f"3) Guardando resultado: {args.salida}")
    save_output(df_out, args.salida)
    print("Listo.")


if __name__ == "__main__":
    main()
