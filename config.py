"""
===========================================================
CONFIGURACIÓN GENERAL DEL SISTEMA
===========================================================
"""

from pathlib import Path

# =========================================================
# INFORMACIÓN GENERAL
# =========================================================

APP_NAME = "Sistema de Captura del Indicador"

VERSION = "1.0.0"

# =========================================================
# DIRECTORIOS
# =========================================================

ROOT = Path(__file__).resolve().parent

DATA_DIR = ROOT / "datos"

ASSETS_DIR = ROOT / "assets"

LOG_DIR = ROOT / "logs"

EXPORT_DIR = ROOT / "exportaciones"

BACKUP_DIR = DATA_DIR / "backups"


# =========================================================
# ARCHIVOS
# =========================================================

MAESTRO = DATA_DIR / "Maestro.xlsx"

CAPTURAS = DATA_DIR / "Capturas.csv"

LOG_FILE = LOG_DIR / "app.log"


# =========================================================
# COLUMNAS DEL MAESTRO
# (Cambiar únicamente si el nombre cambia)
# =========================================================

COL_REGION = "REGION"

COL_EMPRESA = "EMPRESA"

COL_N = "N"

COL_SECCION = "SECCION"

COL_DIVISION = "DIVISION"

COL_GRUPO = "GRUPO"

COL_PONDERADOR = "PONDERADOR"


# =========================================================
# FORMATO FECHAS
# =========================================================

FORMATO_MES = "%b-%y"

FORMATO_FECHA = "%Y-%m-%d %H:%M:%S"


# =========================================================
# INTERFAZ
# =========================================================

PAGE_TITLE = "Sistema de Captura"

PAGE_ICON = "📊"

LAYOUT = "wide"


# =========================================================
# COLORES
# =========================================================

PRIMARY_COLOR = "#005DAA"

SUCCESS_COLOR = "#28A745"

WARNING_COLOR = "#FFC107"

ERROR_COLOR = "#DC3545"


# =========================================================
# MESES
# =========================================================

MESES = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre"
]


# =========================================================
# AÑOS
# =========================================================

ANIOS = list(range(2010, 2051))


# =========================================================
# CREAR DIRECTORIOS SI NO EXISTEN
# =========================================================

for carpeta in [
    DATA_DIR,
    ASSETS_DIR,
    LOG_DIR,
    EXPORT_DIR,
    BACKUP_DIR
]:
    carpeta.mkdir(parents=True, exist_ok=True)
