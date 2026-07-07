"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    settings.py

Descripción:
    Configuración global del proyecto.

    Este archivo contiene todas las rutas y parámetros configurables
    utilizados por el backend.

    Ningún otro módulo debe contener rutas hardcodeadas.
"""

from pathlib import Path


# ==========================================================
# PROJECT
# ==========================================================

ROOT_DIR = Path(__file__).resolve().parents[3]


# ==========================================================
# STORAGE
# ==========================================================

STORAGE_DIR = ROOT_DIR / "storage"

INPUT_DIR = STORAGE_DIR / "input"
OUTPUT_DIR = STORAGE_DIR / "output"
BACKUP_DIR = STORAGE_DIR / "backup"
LOGS_DIR = STORAGE_DIR / "logs"
TEMP_DIR = STORAGE_DIR / "temp"


# ==========================================================
# PROCESS DATE
# ==========================================================

YEAR = "2026"
MONTH = "JULIO"
DAY = "1"


# ==========================================================
# INPUT
# ==========================================================

INPUT_EXCELS_DIR = (
    INPUT_DIR
    / "excels"
    / YEAR
    / MONTH
)

INPUT_PDFS_DIR = (
    INPUT_DIR
    / "pdfs"
    / YEAR
    / MONTH
    / DAY
)


# ==========================================================
# OUTPUT
# ==========================================================

OUTPUT_DIR_DAY = (
    OUTPUT_DIR
    / YEAR
    / MONTH
    / DAY
)


# ==========================================================
# BACKUP
# ==========================================================

BACKUP_DIR_DAY = (
    BACKUP_DIR
    / YEAR
    / MONTH
    / DAY
)


# ==========================================================
# LOGS
# ==========================================================

LOGS_DIR_MONTH = (
    LOGS_DIR
    / YEAR
    / MONTH
)

LOG_FILE = LOGS_DIR_MONTH / f"{DAY}.log"


# ==========================================================
# CREATE DIRECTORIES
# ==========================================================

OUTPUT_DIR_DAY.mkdir(parents=True, exist_ok=True)
BACKUP_DIR_DAY.mkdir(parents=True, exist_ok=True)
LOGS_DIR_MONTH.mkdir(parents=True, exist_ok=True)
TEMP_DIR.mkdir(parents=True, exist_ok=True)