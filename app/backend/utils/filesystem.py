"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    filesystem.py

Descripción:
    Utilidades para trabajar con el sistema de archivos.

    Este módulo centraliza todas las operaciones relacionadas con archivos
    y carpetas utilizadas por el proyecto.
"""

from pathlib import Path


def get_pdf_files(directory: Path) -> list[Path]:
    """
    Devuelve todos los archivos PDF contenidos en un directorio.

    Args:
        directory: Carpeta donde buscar.

    Returns:
        Lista ordenada de archivos PDF.
    """

    return sorted(directory.glob("*.pdf"))


def get_excel_files(directory: Path) -> list[Path]:
    """
    Devuelve todos los archivos Excel contenidos en un directorio.

    Args:
        directory: Carpeta donde buscar.

    Returns:
        Lista ordenada de archivos Excel.
    """

    return sorted(directory.glob("*.xlsx"))


def ensure_directory(directory: Path) -> None:
    """
    Crea una carpeta si no existe.

    Args:
        directory: Carpeta a crear.
    """

    directory.mkdir(parents=True, exist_ok=True)