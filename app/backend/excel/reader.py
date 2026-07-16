"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    reader.py

Descripción:
    Lector de archivos Excel.

    Este módulo únicamente abre un archivo Excel y devuelve el libro y la
    hoja de trabajo principal.
"""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from config.constants import EXCEL_SHEET_NAME


class ExcelReader:
    """
    Lector de documentos Excel.
    """

    def read(self, excel_path: Path) -> tuple[Workbook, Worksheet]:
        """
        Abre un archivo Excel y devuelve el libro junto con la hoja
        'extraccion'.

        Args:
            excel_path: Ruta del archivo Excel.

        Returns:
            Tupla formada por:
                - Workbook
                - Worksheet
        """

        if not excel_path.exists():
            raise FileNotFoundError(excel_path)

        workbook = load_workbook(filename=excel_path)

        if EXCEL_SHEET_NAME not in workbook.sheetnames:
            raise ValueError(f"No existe la hoja '{EXCEL_SHEET_NAME}'.")

        worksheet = workbook[EXCEL_SHEET_NAME]

        return workbook, worksheet
