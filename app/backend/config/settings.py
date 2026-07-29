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

# backend/
BACKEND_DIR = Path(__file__).resolve().parents[1]

# app/
APP_DIR = BACKEND_DIR.parent

# raíz del proyecto
ROOT_DIR = APP_DIR.parent

# ==========================================================
# STORAGE
# ==========================================================

STORAGE_DIR = ROOT_DIR / "storage"

INPUT_DIR = STORAGE_DIR / "input"

# Excel mensuales generados y actualizados por el programa.
MONTHLY_EXCELS_DIR = INPUT_DIR / "excels"

# Plantillas originales de cada punto de venta.
TEMPLATES_DIR = STORAGE_DIR / "templates"

# Copias de seguridad de los Excel modificados.
BACKUP_DIR = STORAGE_DIR / "backup"

# Archivos de registro del programa.
LOGS_DIR = STORAGE_DIR / "logs"

# Registry de entregas sincronizadas.
REGISTRY_DIR = STORAGE_DIR / "registry"
REGISTRY_FILE = REGISTRY_DIR / "imported_deliveries.json"

# ==========================================================
# CREATE DIRECTORIES
# ==========================================================

for directory in (
    MONTHLY_EXCELS_DIR,
    TEMPLATES_DIR,
    BACKUP_DIR,
    LOGS_DIR,
    REGISTRY_DIR,
):
    directory.mkdir(
        parents=True,
        exist_ok=True,
    )
