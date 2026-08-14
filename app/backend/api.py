"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    api.py

Descripción:
    Servidor web local que expone el backend (Importer/Synchronizer)
    a través de una API HTTP y sirve el frontend estático.

    Reutiliza exactamente la misma orquestación que main.py, pero
    recibiendo los Excel de Economato por subida HTTP en vez de por
    el explorador de archivos de tkinter, y devolviendo el resumen
    final como JSON en lugar de imprimirlo por consola.
"""

from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
import shutil
import tempfile

from flask import Flask, jsonify, request, send_file, send_from_directory
from werkzeug.utils import secure_filename

from config.constants import EXCEL_TEMPLATES, MONTHS, SOURCE_EXCEL_EXTENSION
from config.settings import MONTHLY_EXCELS_DIR
from services.excel_template_manager import ExcelTemplateManager
from services.importer import Importer
from services.registry import Registry
from services.synchronizer import Synchronizer
from utils.delivery_identity import build_delivery_key

FRONTEND_DIR = Path(__file__).resolve().parents[1] / "frontend"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB


# ==========================================================
# FRONTEND ESTÁTICO
# ==========================================================


@app.get("/")
def index():
    return send_from_directory(FRONTEND_DIR, "main.html")


@app.get("/<path:filename>")
def static_files(filename: str):
    return send_from_directory(FRONTEND_DIR, filename)


# ==========================================================
# API
# ==========================================================


@app.get("/api/sales-points")
def get_sales_points():
    return jsonify(sorted(EXCEL_TEMPLATES))


@app.get("/api/periods")
def get_periods():
    """
    Devuelve los años y meses para los que ya existe al menos un
    Excel mensual, para rellenar el selector de exportación.
    """

    periods: dict[str, list[str]] = {}

    if MONTHLY_EXCELS_DIR.is_dir():
        for year_dir in sorted(MONTHLY_EXCELS_DIR.iterdir()):
            if not year_dir.is_dir():
                continue

            months_present = [
                month_dir.name
                for month_dir in sorted(
                    year_dir.iterdir(),
                    key=lambda path: MONTHS.index(path.name.upper())
                    if path.name.upper() in MONTHS
                    else 99,
                )
                if month_dir.is_dir() and month_dir.name.upper() in MONTHS
            ]

            if months_present:
                periods[year_dir.name] = months_present

    return jsonify(periods)


@app.post("/api/sync")
def sync():
    uploaded_files = request.files.getlist("files")
    uploaded_files = [file for file in uploaded_files if file.filename]

    if not uploaded_files:
        return jsonify({"error": "No se ha recibido ningún archivo Excel."}), 400

    invalid_files = [
        file.filename
        for file in uploaded_files
        if not file.filename.lower().endswith(SOURCE_EXCEL_EXTENSION)
    ]

    if invalid_files:
        return (
            jsonify(
                {
                    "error": "Solo se admiten archivos "
                    f"{SOURCE_EXCEL_EXTENSION}: {', '.join(invalid_files)}"
                }
            ),
            400,
        )

    temp_dir = Path(tempfile.mkdtemp(prefix="iberostar_sync_"))

    try:
        saved_paths: list[Path] = []

        for index, file in enumerate(uploaded_files):
            safe_name = secure_filename(file.filename) or f"archivo_{index}.xlsx"
            destination = temp_dir / f"{index:03d}_{safe_name}"
            file.save(destination)
            saved_paths.append(destination)

        try:
            with Registry() as registry:
                importer = Importer(registry=registry)
                synchronizer = Synchronizer(registry=registry)

                deliveries = importer.run(excel_files=saved_paths)
                import_summary = importer.last_summary

                known_delivery_keys = {
                    build_delivery_key(delivery) for delivery in deliveries
                }

                recovered_deliveries, pending_warnings = (
                    registry.get_pending_deliveries()
                )
                recovered_deliveries = [
                    delivery
                    for delivery in recovered_deliveries
                    if build_delivery_key(delivery) not in known_delivery_keys
                ]

                deliveries = deliveries + recovered_deliveries

                synchronization_totals = synchronizer.run(deliveries)
        except RuntimeError as error:
            return (
                jsonify(
                    {
                        "error": "El Registry está en uso por otro proceso. "
                        f"Inténtalo de nuevo en unos segundos. Detalle: {error}"
                    }
                ),
                409,
            )
        except Exception as error:  # Error inesperado de programación.
            app.logger.exception("Fallo inesperado durante la sincronización.")
            return (
                jsonify(
                    {
                        "error": "Error inesperado durante la sincronización: "
                        f"{type(error).__name__}: {error}"
                    }
                ),
                500,
            )

        has_incidents = (
            import_summary.has_incidents()
            or bool(pending_warnings)
            or bool(synchronization_totals.error_deliveries)
        )

        return jsonify(
            {
                "has_incidents": has_incidents,
                "import_summary": asdict(import_summary),
                "pending_warnings": pending_warnings,
                "synchronization_totals": asdict(synchronization_totals),
            }
        )

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@app.get("/api/export")
def export_excel():
    sales_point = request.args.get("sales_point", "").strip()
    year_text = request.args.get("year", "").strip()
    month_text = request.args.get("month", "").strip()

    if not sales_point or not year_text or not month_text:
        return (
            jsonify({"error": "Faltan parámetros: sales_point, year, month."}),
            400,
        )

    try:
        year = int(year_text)
        month = int(month_text)
    except ValueError:
        return jsonify({"error": "year y month deben ser números."}), 400

    try:
        excel_path = ExcelTemplateManager().get_excel_path(
            sales_point=sales_point,
            year=year,
            month=month,
        )
    except ValueError as error:
        return jsonify({"error": str(error)}), 400

    if not excel_path.is_file():
        return (
            jsonify(
                {
                    "error": "Todavía no existe un Excel sincronizado para "
                    f"{sales_point} en {month_text}/{year_text}."
                }
            ),
            404,
        )

    return send_file(excel_path, as_attachment=True, download_name=excel_path.name)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
