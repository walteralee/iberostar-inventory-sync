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

        if current_value in (
            None,
            "",
        ):
            current_value = 0

        print(f"{'':18}│ " f"Fila destino            : {row}")

        print(f"{'':18}│ " f"Columna destino         : {column}")

        print(f"{'':18}│ " f"Valor anterior          : {current_value}")

        print(f"{'':18}│ " f"Cantidad añadida        : {quantity}")

        cell.value = current_value + quantity

        print(f"{'':18}│ " f"Valor final             : {cell.value}")

        print(f"{'':18}│ " f"Resultado               : CANTIDAD ACUMULADA")
