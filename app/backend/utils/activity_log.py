"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    activity_log.py

Descripción:
    Registro persistente de incidencias de importación en
    storage/logs/, para no depender únicamente de la consola.
"""

from __future__ import annotations

from datetime import datetime, timezone

from config.settings import LOGS_DIR

_LOG_FILE = LOGS_DIR / "importacion_incidencias.log"


def log_incident(message: str) -> None:
    """
    Añade una línea con timestamp UTC al log de incidencias.
    """

    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")

    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    with _LOG_FILE.open("a", encoding="utf-8") as file:
        file.write(f"[{timestamp}] {message}\n")
