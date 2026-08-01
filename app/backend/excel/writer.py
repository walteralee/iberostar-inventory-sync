"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    writer.py

Descripción:
    Servicio encargado de escribir cantidades en los Excel
    mensuales.
"""

from math import isfinite

from openpyxl.worksheet.worksheet import Worksheet


class ExcelWriter:
    """
    Servicio encargado de escribir cantidades en un Excel.
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

        Si la celda ya contiene un valor numérico,
        ambas cantidades se acumulan.

        Args:
            worksheet: Hoja donde se escribirá.
            row: Fila destino.
            column: Columna destino.
            quantity: Cantidad que debe añadirse.

        Raises:
            ValueError:
                Si alguno de los datos no es válido.
        """

        self._validate_arguments(
            row=row,
            column=column,
            quantity=quantity,
        )

        cell = worksheet.cell(
            row=row,
            column=column,
        )

        current_value = cell.value

        if current_value in (None, ""):
            current_value = 0

        current_value = self._validate_numeric_value(
            current_value,
            f"La celda {cell.coordinate}",
        )

        cell.value = current_value + quantity

    def _validate_arguments(
        self,
        row: int,
        column: int,
        quantity: float,
    ) -> None:
        """
        Valida los argumentos de escritura.
        """

        if not isinstance(row, int) or row < 1:
            raise ValueError(f"Fila inválida: {row}")

        if not isinstance(column, int) or column < 1:
            raise ValueError(f"Columna inválida: {column}")

        self._validate_numeric_value(
            quantity,
            "La cantidad",
        )

    def _validate_numeric_value(
        self,
        value: object,
        field_name: str,
    ) -> float:
        """
        Valida un valor numérico.

        Rechaza:

        - None
        - bool
        - texto
        - NaN
        - infinito
        """

        if isinstance(value, bool) or not isinstance(
            value,
            (int, float),
        ):
            raise ValueError(f"{field_name} debe ser numérica.")

        value = float(value)

        if not isfinite(value):
            raise ValueError(f"{field_name} contiene un número no válido.")

        return value
