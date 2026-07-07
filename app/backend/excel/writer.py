"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    writer.py

Descripción:
    Escritura de cantidades dentro del Excel.
"""

from openpyxl.worksheet.worksheet import Worksheet


class ExcelWriter:

    def write(
        self,
        worksheet: Worksheet,
        row: int,
        column: int,
        quantity: float
    ) -> None:
        """
        Escribe una cantidad en una celda.
        """

        worksheet.cell(
            row=row,
            column=column
        ).value = quantity