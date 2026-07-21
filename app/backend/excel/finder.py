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

    def build_product_index(
        self,
        worksheet: Worksheet,
    ) -> dict[str, int]:
        """
        Construye un índice que relaciona cada código de producto
        con su fila correspondiente.
        """

        print()
        print("-" * 100)
        print("CONSTRUCCIÓN DEL ÍNDICE DE PRODUCTOS")
        print("-" * 100)

        index: dict[str, int] = {}

        for row in range(
            FIRST_PRODUCT_ROW,
            worksheet.max_row + 1,
        ):
            value = worksheet.cell(
                row=row,
                column=PRODUCT_CODE_COLUMN,
            ).value

            if value is None:
                continue

            code = str(value).strip()

            if not code:
                continue

            index[code] = row

        print(f"Productos indexados : {len(index)}")
        print(f"Primera fila        : {FIRST_PRODUCT_ROW}")
        print(f"Última fila leída   : {worksheet.max_row}")
        print("Estado              : ÍNDICE CONSTRUIDO")
        print("-" * 100)

        return index

    def find_day_column(
        self,
        day: int,
    ) -> int:
        """
        Calcula la columna correspondiente al día del mes.
        """

        print()
        print("-" * 100)
        print("LOCALIZACIÓN DE LA COLUMNA DEL DÍA")
        print("-" * 100)

        if not 1 <= day <= 31:
            raise ValueError(f"Día inválido: {day}")

        column = FIRST_DAY_COLUMN + (day - 1)

        print(f"Día del movimiento : {day}")
        print(f"Fila cabecera      : {DAY_HEADER_ROW}")
        print(f"Columna Excel      : {column}")
        print("Estado             : COLUMNA LOCALIZADA")
        print("-" * 100)

        return column
