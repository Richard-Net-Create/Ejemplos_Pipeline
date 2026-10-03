"""
CONSOLIDADOR DE REPORTES DE VENTAS
==================================
Une todos los .xlsx de una carpeta, elimina duplicados, normaliza fechas,
genera reportes de ventas y (opcional) los envía por correo.

Uso:
    python juntar_excels.py                  # solo genera el Excel
    python juntar_excels.py --enviar-correo  # además lo envía por Gmail

Variables de entorno para el correo (ver .env.example):
    GMAIL_USER, GMAIL_APP_PASSWORD, MAIL_TO
"""

import argparse
import glob
import os
from datetime import datetime

import pandas as pd

META = 50000
SALIDA = "Reporte_Maestro_PRO.xlsx"


def cargar_archivos(carpeta: str) -> pd.DataFrame:
    frames = []
    for archivo in glob.glob(os.path.join(carpeta, "*.xlsx")):
        if "Reporte_Maestro" in archivo:
            continue
        df = pd.read_excel(archivo)
        df["Origen"] = os.path.splitext(os.path.basename(archivo))[0]
        frames.append(df)
    if not frames:
        raise SystemExit(f"No se encontraron archivos .xlsx en {carpeta!r}")
    return pd.concat(frames, ignore_index=True).drop_duplicates()


def construir_reportes(df: pd.DataFrame) -> dict:
    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"])
        df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)

    total = df["Venta_S"].sum()

    por_vendedor = (
        df.groupby("Vendedor")["Venta_S"].sum().reset_index()
        .sort_values("Venta_S", ascending=False)
    )
    por_vendedor["% Participación"] = (por_vendedor["Venta_S"] / total * 100).round(2)
    por_vendedor.columns = ["Vendedor", "Total Venta_S", "% Participación"]

    alerta = por_vendedor.copy()
    alerta[f"Meta S/{META}"] = META
    alerta["Diferencia"] = alerta["Total Venta_S"] - META
    alerta["Estado"] = alerta["Diferencia"].apply(
        lambda x: "CUMPLIÓ" if x >= 0 else "NO CUMPLIÓ"
    )

    if "Mes" in df.columns:
        por_mes = df.groupby("Mes")["Venta_S"].sum().reset_index()
        por_mes.columns = ["Mes", "Total Venta_S"]
    else:
        por_mes = pd.DataFrame({"Info": ["No hay columna Fecha"]})

    def agrupar(col):
        return (
            df.groupby(col)["Venta_S"].sum().reset_index()
            .sort_values("Venta_S", ascending=False)
        )

    return {
        "total": total,
        "hojas": {
            "Datos_Completos": df,
            "Por_Vendedor_%": por_vendedor,
            "Alerta_Meta": alerta,
            "Por_Mes": por_mes,
            "Por_Tipo_Celular": agrupar("Tipo_Celular"),
            "Por_Canal": agrupar("Canal_Venta"),
            "Por_Metodo_Pago": agrupar("Metodo_Pago"),
        },
    }


def guardar_excel(hojas: dict, ruta: str) -> None:
    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        for nombre, hoja in hojas.items():
            hoja.to_excel(writer, sheet_name=nombre, index=False)


def enviar_correo(adjunto: str) -> None:
    import yagmail  # se importa aquí para que sea opcional

    usuario = os.environ.get("GMAIL_USER")
    clave = os.environ.get("GMAIL_APP_PASSWORD")
    destino = os.environ.get("MAIL_TO", usuario)
    if not usuario or not clave:
        raise SystemExit("Faltan GMAIL_USER y/o GMAIL_APP_PASSWORD en el entorno.")

    asunto = f"Reporte de ventas - {datetime.now():%d/%m/%Y}"
    cuerpo = (
        "Hola equipo,\n\nAdjunto el Reporte Maestro con corte al día de hoy.\n"
        "Incluye: ranking por vendedor, ventas por mes, alerta de meta y más.\n\n"
        "Saludos,\nRobot de Reportes"
    )
    yagmail.SMTP(usuario, clave).send(
        to=destino, subject=asunto, contents=cuerpo, attachments=adjunto
    )
    print(f"Correo enviado a {destino}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--carpeta", default=".", help="Carpeta con los .xlsx")
    parser.add_argument("--enviar-correo", action="store_true")
    args = parser.parse_args()

    df = cargar_archivos(args.carpeta)
    resultado = construir_reportes(df)
    guardar_excel(resultado["hojas"], SALIDA)
    print(f"Listo. Total general: S/{resultado['total']:,.2f}")
    print(f"Se creó {SALIDA} con {len(resultado['hojas'])} pestañas")

    if args.enviar_correo:
        enviar_correo(SALIDA)


if __name__ == "__main__":
    main()
