"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    settings.py

Descripción:
    Configuración global del proyecto.

    Este archivo contiene todas las rutas utilizadas por el backend.

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

INPUT_EXCELS_DIR = INPUT_DIR / "excels"
INPUT_PDFS_DIR = INPUT_DIR / "pdfs"

OUTPUT_DIR = STORAGE_DIR / "output"
BACKUP_DIR = STORAGE_DIR / "backup"
LOGS_DIR = STORAGE_DIR / "logs"
TEMP_DIR = STORAGE_DIR / "temp"

TEMPLATES_DIR = STORAGE_DIR / "templates"

REGISTRY_DIR = STORAGE_DIR / "registry"
REGISTRY_FILE = REGISTRY_DIR / "imported_deliveries.json"


# ==========================================================
# CREATE DIRECTORIES
# ==========================================================

INPUT_EXCELS_DIR.mkdir(parents=True, exist_ok=True)
INPUT_PDFS_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)
TEMP_DIR.mkdir(parents=True, exist_ok=True)

TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
REGISTRY_DIR.mkdir(parents=True, exist_ok=True)
