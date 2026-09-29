"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    api.py

Descripción:
    Servidor web local que expone el backend a través de una API HTTP
    y sirve el frontend estático.

    Solo escucha en 127.0.0.1: la aplicación de escritorio (desktop.py)
    lo arranca en segundo plano y lo muestra dentro de su propia
    ventana. Todas las respuestas de error de /api son JSON con una
    clave ``error`` legible por el usuario final.
"""

from __future__ import annotations

from pathlib import Path
import logging
import os
import shutil
import subprocess
import sys
import tempfile

from flask import Flask, jsonify, request, send_file, send_from_directory
from werkzeug.exceptions import HTTPException
from werkzeug.utils import secure_filename

from config.constants import (
    EXCEL_TEMPLATES,
    MONTHS,
    PROJECT_NAME,
    PROJECT_VERSION,
    SOURCE_EXCEL_EXTENSION,
)
from config.settings import DATA_DIR, FRONTEND_DIR, MONTHLY_EXCELS_DIR
from services.excel_template_manager import ExcelTemplateManager
from services.sync_pipeline import run_sync

logger = logging.getLogger(__name__)

MAX_UPLOAD_BYTES = 200 * 1024 * 1024  # 200 MB


class ExportNotFoundError(LookupError):
    """No existe todavía un Excel sincronizado para el periodo pedido."""


def resolve_export_path(sales_point: str, year: int, month: int) -> Path:
    """
    Ruta del Excel mensual ya sincronizado de un punto de venta.

    Raises:
        ValueError: parámetros inválidos.
        ExportNotFoundError: el Excel todavía no existe.
    """

    if sales_point not in EXCEL_TEMPLATES:
        raise ValueError(f"Punto de venta desconocido: {sales_point}.")

    excel_path = ExcelTemplateManager().get_excel_path(
        sales_point=sales_point,
        year=year,
        month=month,
    )

    if not excel_path.is_file():
        raise ExportNotFoundError(
            "Todavía no existe un Excel sincronizado para "
            f"{sales_point.replace('_', ' ')} en {month:02d}/{year}."
        )

    return excel_path


def list_periods() -> dict[str, list[str]]:
    """
    Años y meses para los que ya existe al menos un Excel mensual.
    """

    periods: dict[str, list[str]] = {}

    if not MONTHLY_EXCELS_DIR.is_dir():
        return periods

    for year_dir in sorted(MONTHLY_EXCELS_DIR.iterdir()):
        if not year_dir.is_dir():
            continue

        months_present = sorted(
            (
                month_dir.name
                for month_dir in year_dir.iterdir()
                if month_dir.is_dir() and month_dir.name.upper() in MONTHS
            ),
            key=lambda name: MONTHS.index(name.upper()),
        )

        if months_present:
            periods[year_dir.name] = months_present

    return periods


def open_in_file_explorer(path: Path) -> None:
    """Abre una carpeta con el explorador de archivos del sistema."""

    path.mkdir(parents=True, exist_ok=True)

    if os.name == "nt":
        os.startfile(path)  # noqa: S606 - ruta interna, no controlada por el usuario
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


def _error(message: str, status: int):
    return jsonify({"error": message}), status


def create_app() -> Flask:
    app = Flask(__name__, static_folder=None)
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

    # ======================================================
    # ERRORES
    # ======================================================

    @app.errorhandler(HTTPException)
    def handle_http_error(error: HTTPException):
        if not request.path.startswith("/api/"):
            return error

        messages = {
            404: "Recurso no encontrado.",
            405: "Operación no permitida.",
            413: "Los archivos superan el tamaño máximo permitido (200 MB).",
        }

        return _error(messages.get(error.code, error.description), error.code)

    @app.errorhandler(Exception)
    def handle_unexpected_error(error: Exception):
        logger.exception("Error inesperado en %s", request.path)

        return _error(
            "Se ha producido un error inesperado. Los detalles se han "
            "guardado en el registro de la aplicación.",
            500,
        )

    # ======================================================
    # API
    # ======================================================

    @app.get("/api/info")
    def get_info():
        return jsonify(
            {
                "name": PROJECT_NAME,
                "version": PROJECT_VERSION,
                "data_dir": str(DATA_DIR),
            }
        )

    @app.get("/api/ping")
    def ping():
        return jsonify({"ok": True})

    @app.get("/api/sales-points")
    def get_sales_points():
        return jsonify(sorted(EXCEL_TEMPLATES))

    @app.get("/api/periods")
    def get_periods():
        return jsonify(list_periods())

    @app.post("/api/open-folder")
    def open_folder():
        open_in_file_explorer(MONTHLY_EXCELS_DIR)
        return jsonify({"ok": True})

    @app.post("/api/sync")
    def sync():
        uploaded_files = [
            file for file in request.files.getlist("files") if file.filename
        ]

        if not uploaded_files:
            return _error("No se ha recibido ningún archivo Excel.", 400)

        invalid_files = [
            file.filename
            for file in uploaded_files
            if not file.filename.lower().endswith(SOURCE_EXCEL_EXTENSION)
        ]

        if invalid_files:
            return _error(
                f"Solo se admiten archivos {SOURCE_EXCEL_EXTENSION}: "
                f"{', '.join(invalid_files)}",
                400,
            )

        temp_dir = Path(tempfile.mkdtemp(prefix="iberostar_sync_"))

        try:
            saved_paths: list[Path] = []

            for index, file in enumerate(uploaded_files):
                safe_name = secure_filename(file.filename) or "archivo.xlsx"
                destination = temp_dir / f"{index:03d}_{safe_name}"
                file.save(destination)
                saved_paths.append(destination)

            logger.info("Sincronizando %d archivo(s).", len(saved_paths))

            try:
                result = run_sync(saved_paths)
            except RuntimeError as error:
                logger.warning("Registry bloqueado: %s", error)
                return _error(
                    "Ya hay una sincronización en curso. Espera a que "
                    "termine e inténtalo de nuevo.",
                    409,
                )

            totals = result.synchronization_totals
            logger.info(
                "Sincronización terminada: %d entregas, %d errores.",
                totals.synchronized_deliveries,
                totals.error_deliveries,
            )

            return jsonify(result.to_dict())

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    @app.get("/api/export")
    def export_excel():
        sales_point = request.args.get("sales_point", "").strip()

        try:
            year = int(request.args.get("year", ""))
            month = int(request.args.get("month", ""))
        except ValueError:
            return _error("Selecciona un año y un mes válidos.", 400)

        try:
            excel_path = resolve_export_path(sales_point, year, month)
        except ValueError as error:
            return _error(str(error), 400)
        except ExportNotFoundError as error:
            return _error(str(error), 404)

        return send_file(
            excel_path,
            as_attachment=True,
            download_name=excel_path.name,
            max_age=0,
        )

    # ======================================================
    # FRONTEND ESTÁTICO
    # ======================================================

    @app.get("/")
    def index():
        return send_from_directory(FRONTEND_DIR, "main.html", max_age=0)

    @app.get("/<path:filename>")
    def static_files(filename: str):
        if filename.startswith("api/"):
            return _error("Recurso no encontrado.", 404)

        return send_from_directory(FRONTEND_DIR, filename, max_age=0)

    return app


if __name__ == "__main__":
    # Modo desarrollo: servidor solo, abrir http://127.0.0.1:5000/
    logging.basicConfig(level=logging.INFO)
    create_app().run(host="127.0.0.1", port=5000, debug=False)
