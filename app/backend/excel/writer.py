"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    writer.py

Descripción:
    Servicio encargado de escribir cantidades en los Excel
    mensuales.
"""

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
            ValueError: Si la fila, la columna, la cantidad
                o el valor actual de la celda no son válidos.
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

        if current_value in (
            None,
            "",
        ):
            current_value = 0

        if not isinstance(
            current_value,
            (
                int,
                float,
            ),
        ):
            raise ValueError(
                f"La celda {cell.coordinate} contiene "
                f"un valor no numérico: {current_value!r}"
            )

        print(f"{'':18}│ Fila destino            : {row}")
        print(f"{'':18}│ Columna destino         : {column}")
        print(f"{'':18}│ Valor anterior          : {current_value}")
        print(f"{'':18}│ Cantidad añadida        : {quantity}")

        cell.value = current_value + quantity

        print(f"{'':18}│ Valor final             : {cell.value}")
        print(f"{'':18}│ Resultado               : CANTIDAD ACUMULADA")

    def _validate_arguments(
        self,
        row: int,
        column: int,
        quantity: float,
    ) -> None:
        """
        Valida los argumentos de escritura.

        Args:
            row: Fila destino.
            column: Columna destino.
            quantity: Cantidad que debe escribirse.

        Raises:
            ValueError: Si alguno de los argumentos no es válido.
        """

        if not isinstance(row, int) or row < 1:
            raise ValueError(f"Fila inválida: {row}")

        if not isinstance(column, int) or column < 1:
            raise ValueError(f"Columna inválida: {column}")

        if isinstance(quantity, bool) or not isinstance(
            quantity,
            (
                int,
                float,
            ),
        ):
            raise ValueError("La cantidad debe ser numérica.")
