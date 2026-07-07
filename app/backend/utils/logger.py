"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    logger.py

Descripción:
    Configuración del sistema de logs del proyecto.

    Centraliza toda la salida de información, advertencias y errores de la
    aplicación.
"""

import logging
from pathlib import Path


def create_logger(log_file: Path) -> logging.Logger:
    """
    Crea y configura el logger principal del proyecto.

    Args:
        log_file: Ruta donde se almacenará el archivo de log.

    Returns:
        Logger configurado.
    """

    logger = logging.getLogger("iberostar")

    if logger.hasHandlers():
        return logger

    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "[%(asctime)s] %(levelname)s - %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S",
    )

    # Consola
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Archivo
    log_file.parent.mkdir(parents=True, exist_ok=True)

    file_handler = logging.FileHandler(
        log_file,
        encoding="utf-8",
    )

    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger