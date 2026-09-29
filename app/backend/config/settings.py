"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    settings.py

Descripción:
    Configuración global del proyecto.

    Este archivo contiene todas las rutas utilizadas por el backend.
    Ningún otro módulo debe contener rutas hardcodeadas.

    Se distinguen dos tipos de rutas:

    - Recursos (solo lectura): frontend y plantillas originales. En
      desarrollo están dentro del repositorio; en la aplicación
      empaquetada con PyInstaller están dentro del propio ejecutable.

    - Datos (lectura/escritura): Excel mensuales, plantillas en uso,
      copias de seguridad, Registry y logs. En desarrollo se guardan en
      ``storage/``; en la aplicación instalada se guardan en la carpeta
      Documentos del usuario, porque la carpeta de instalación
      (Program Files) no admite escritura.

    La variable de entorno ``IBEROSTAR_DATA_DIR`` permite forzar
    cualquier otra carpeta de datos.
"""

from __future__ import annotations

from pathlib import Path
import os
import shutil
import sys

APP_DISPLAY_NAME = "Iberostar Gestor de Pedidos"

# ==========================================================
# RESOURCES (READ-ONLY)
# ==========================================================

IS_FROZEN = getattr(sys, "frozen", False)

if IS_FROZEN:
    # Carpeta donde PyInstaller descomprime los recursos empaquetados.
    RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
else:
    # Raíz del repositorio (config/ -> backend/ -> app/ -> raíz).
    RESOURCE_DIR = Path(__file__).resolve().parents[3]

FRONTEND_DIR = RESOURCE_DIR / "app" / "frontend"

# Plantillas originales que acompañan a cada versión del programa.
BUNDLED_TEMPLATES_DIR = RESOURCE_DIR / "storage" / "templates"


# ==========================================================
# DATA (READ-WRITE)
# ==========================================================


def _documents_dir() -> Path:
    """
    Carpeta Documentos real del usuario, respetando redirecciones
    (por ejemplo, OneDrive) en Windows.
    """

    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            # FOLDERID_Documents
            folder_id = ctypes.c_char_p(
                bytes.fromhex("D0 9A D3 FD 8F 23 AF 46 AD B4 6C 85 48 03 69 C7")
            )
            path_pointer = ctypes.c_wchar_p()

            result = ctypes.windll.shell32.SHGetKnownFolderPath(
                folder_id,
                0,
                None,
                ctypes.byref(path_pointer),
            )

            if result == 0 and path_pointer.value:
                documents = Path(path_pointer.value)
                ctypes.windll.ole32.CoTaskMemFree(
                    ctypes.cast(path_pointer, wintypes.LPVOID)
                )
                return documents

        except (AttributeError, OSError):
            pass

    return Path.home() / "Documents"


def _resolve_data_dir() -> Path:
    override = os.environ.get("IBEROSTAR_DATA_DIR", "").strip()

    if override:
        return Path(override).expanduser()

    if IS_FROZEN:
        return _documents_dir() / APP_DISPLAY_NAME

    return RESOURCE_DIR / "storage"


DATA_DIR = _resolve_data_dir()

if IS_FROZEN:
    # Nombres pensados para el usuario final, que verá esta carpeta.
    MONTHLY_EXCELS_DIR = DATA_DIR / "Excel mensuales"
    TEMPLATES_DIR = DATA_DIR / "Plantillas"
    BACKUP_DIR = DATA_DIR / "Copias de seguridad"
    LOGS_DIR = DATA_DIR / "Sistema" / "logs"
    REGISTRY_DIR = DATA_DIR / "Sistema" / "registry"
else:
    # Estructura histórica del repositorio.
    MONTHLY_EXCELS_DIR = DATA_DIR / "input" / "excels"
    TEMPLATES_DIR = DATA_DIR / "templates"
    BACKUP_DIR = DATA_DIR / "backup"
    LOGS_DIR = DATA_DIR / "logs"
    REGISTRY_DIR = DATA_DIR / "registry"

REGISTRY_FILE = REGISTRY_DIR / "imported_deliveries.json"


# ==========================================================
# CREATE DIRECTORIES
# ==========================================================


def _prepare_data_dir() -> None:
    """
    Crea la estructura de datos y copia las plantillas originales que
    falten. Las plantillas existentes nunca se sobrescriben, porque el
    programa les añade productos nuevos con el uso.
    """

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

    if TEMPLATES_DIR.resolve() == BUNDLED_TEMPLATES_DIR.resolve():
        return

    if not BUNDLED_TEMPLATES_DIR.is_dir():
        return

    for template in BUNDLED_TEMPLATES_DIR.glob("*.xlsx"):
        destination = TEMPLATES_DIR / template.name

        if not destination.exists():
            shutil.copy2(template, destination)


_prepare_data_dir()
