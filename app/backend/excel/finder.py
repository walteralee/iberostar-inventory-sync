"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    finder.py

Descripción:
    Utilidades para localizar información dentro del Excel.

    Este módulo permite localizar la fila de un producto mediante su código
    de artículo y calcular la columna correspondiente al día del mes.
"""

from openpyxl.worksheet.worksheet import Worksheet

from config.constants import (
    PRODUCT_CODE_COLUMN,
    FIRST_PRODUCT_ROW,
    DAY_HEADER_ROW,
    FIRST_DAY_COLUMN,
)


class ExcelFinder:
    """
    Localizador de información dentro del Excel.
    """

    def build_product_index(self, worksheet: Worksheet) -> dict[int, int]:
        """
        Construye un índice de productos.

        Args:
            worksheet: Hoja del Excel.

        Returns:
            Diccionario:
                codigo_articulo -> fila
        """

        index: dict[int, int] = {}

        for row in range(FIRST_PRODUCT_ROW, worksheet.max_row + 1):

            value = worksheet.cell(
                row=row,
                column=PRODUCT_CODE_COLUMN
            ).value

            if value is None:
                continue

            try:
                code = int(value)
            except (TypeError, ValueError):
                continue

            index[code] = row

        return index

    def find_product(
        self,
        product_index: dict[int, int],
        product_code: int,
    ) -> int | None:
        """
        Devuelve la fila correspondiente a un código de artículo.

        Args:
            product_index: Índice de productos.
            product_code: Código del artículo.

        Returns:
            Número de fila o None si no existe.
        """

        return product_index.get(product_code)

    def find_day_column(self, day: int) -> int:
        """
        Calcula la columna correspondiente al día del mes.

        Args:
            day: Día del mes.

        Returns:
            Número de columna.
        """

        if not 1 <= day <= 31:
            raise ValueError(f"Día inválido: {day}")

        return FIRST_DAY_COLUMN + (day - 1)