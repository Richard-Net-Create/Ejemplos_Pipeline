# Pipeline de datos en Python: Excel → limpieza SQL → PostgreSQL

Scripts en Python para consolidar archivos Excel, limpiarlos con SQL y cargarlos
en PostgreSQL, además de generar reportes de ventas listos para enviar por correo.

## Qué incluye

| Script | Qué hace |
|---|---|
| `juntar_excels.py` | Une todos los `.xlsx` de una carpeta, elimina duplicados, normaliza fechas y genera un Excel de 7 pestañas (participación por vendedor, alerta de meta, ventas por mes, tipo de celular, canal y método de pago). Opcionalmente lo envía por correo. |
| `pipeline.py` | Carga un Excel/CSV como tabla SQLite en memoria, aplica una consulta SQL de limpieza (duplicados, nulos, condiciones) y exporta el resultado a CSV/Excel. |
| `pipeline_postgres.py` | Hace la misma limpieza y carga el resultado en una tabla de PostgreSQL con SQLAlchemy + pg8000. |

## Tecnologías

Python · pandas · SQL (SQLite) · SQLAlchemy · PostgreSQL · openpyxl · yagmail

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Configuración

Copia `.env.example` a `.env`, completa tus valores y cárgalos en tu terminal
(PowerShell: `$env:PG_PASSWORD="..."`; Linux/macOS: `export PG_PASSWORD=...`).
**Nunca subas contraseñas ni el archivo `.env` al repositorio.**

## Uso

```bash
# Consolidar Excels y generar el reporte maestro
python juntar_excels.py --carpeta ./datos

# Lo mismo y enviarlo por correo (requiere GMAIL_USER y GMAIL_APP_PASSWORD)
python juntar_excels.py --carpeta ./datos --enviar-correo

# Limpieza SQL a archivo
python pipeline.py --entrada datos/ejemplo_nomina.xlsx --salida salida/limpio.xlsx

# Limpieza SQL + carga a PostgreSQL
python pipeline_postgres.py
```

Cada script imprime cuántas filas entraron, cuántas salieron y cuántas se
descartaron, para poder auditar el resultado.

## Datos

Los archivos reales no se incluyen. Para probar el proyecto, usa un Excel de
ejemplo con datos ficticios dentro de `datos/` (nombre `ejemplo_*.xlsx`).
