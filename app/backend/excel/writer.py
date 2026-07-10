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
    """
    Servicio encargado de escribir cantidades en el Excel.
    """

    def write(
        self,
        worksheet: Worksheet,
        row: int,
        column: int,
        quantity: float,
    ) -> None:
        """
        Escribe una cantidad en una celda.

        Si la celda ya contiene una cantidad, ambas se
        acumulan.
        """

        cell = worksheet.cell(
            row=row,
            column=column,
        )

        current_value = cell.value

        if current_value in (None, ""):
            current_value = 0

        cell.value = current_value + quantity
