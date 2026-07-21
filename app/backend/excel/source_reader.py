"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    source_reader.py

Descripción:
    Servicio encargado de abrir y proporcionar acceso
    al Excel de origen.
"""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


class SourceReader:
    """
    Servicio encargado de abrir el Excel de origen.
    """

    def read(
        self,
        excel_file: Path,
    ) -> Worksheet:
        """
        Abre el Excel y devuelve la hoja activa.

        Args:
            excel_file: Ruta del Excel de origen.

        Returns:
            Hoja activa del Excel.
        """

        workbook = load_workbook(
            filename=excel_file,
            data_only=True,
        )

        return workbook.active
