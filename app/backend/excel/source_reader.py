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

from config.constants import (
    SOURCE_HEADER_ROW,
    SOURCE_DATE_COLUMN,
    SOURCE_SALES_POINT_COLUMN,
    SOURCE_GROUP_COLUMN,
    SOURCE_PRODUCT_CODE_COLUMN,
    SOURCE_PRODUCT_NAME_COLUMN,
    SOURCE_FORMAT_COLUMN,
    SOURCE_QUANTITY_COLUMN,
    SOURCE_PRICE_COLUMN,
)


class SourceReader:
    """
    Servicio encargado de abrir el Excel de origen.

    No depende de la hoja activa del libro. Busca
    automáticamente una hoja que contenga la estructura
    mínima necesaria para que el Importer pueda procesarla.
    """

    _REQUIRED_COLUMNS = (
        SOURCE_DATE_COLUMN,
        SOURCE_SALES_POINT_COLUMN,
        SOURCE_GROUP_COLUMN,
        SOURCE_PRODUCT_CODE_COLUMN,
        SOURCE_PRODUCT_NAME_COLUMN,
        SOURCE_FORMAT_COLUMN,
        SOURCE_QUANTITY_COLUMN,
        SOURCE_PRICE_COLUMN,
    )

    def read(
        self,
        excel_file: Path,
    ) -> Worksheet:
        """
        Abre el Excel y devuelve la hoja de movimientos.

        Args:
            excel_file: Ruta del Excel de origen.

        Returns:
            Hoja válida del Excel.

        Raises:
            ValueError:
                Si no existe ninguna hoja con la estructura mínima.
        """

        workbook = load_workbook(
            filename=excel_file,
            data_only=True,
        )

        for worksheet in workbook.worksheets:
            if self._is_valid_source_sheet(worksheet):
                return worksheet

        workbook.close()

        raise ValueError(
            "No se encontró ninguna hoja con la estructura mínima "
            "necesaria para importar los movimientos de Economato."
        )

    # ======================================================
    # PRIVATE
    # ======================================================

    def _is_valid_source_sheet(
        self,
        worksheet: Worksheet,
    ) -> bool:
        """
        Comprueba si una hoja contiene las columnas y filas
        mínimas necesarias para que pueda procesarla el Importer.

        No compara el texto exacto de las cabeceras, ya que los
        informes reales pueden utilizar nombres diferentes.
        """

        required_last_column = max(self._REQUIRED_COLUMNS)

        if worksheet.max_column < required_last_column:
            return False

        if worksheet.max_row <= SOURCE_HEADER_ROW:
            return False

        return all(
            not self._is_blank(
                worksheet.cell(
                    row=SOURCE_HEADER_ROW,
                    column=column,
                ).value
            )
            for column in self._REQUIRED_COLUMNS
        )

    def _is_blank(
        self,
        value: object,
    ) -> bool:
        """
        Comprueba si una celda está vacía o contiene
        únicamente espacios.
        """

        return value is None or str(value).strip() == ""
