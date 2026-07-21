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

MONTHLY_EXCELS_DIR.mkdir(parents=True, exist_ok=True)

TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)

REGISTRY_DIR.mkdir(parents=True, exist_ok=True)
