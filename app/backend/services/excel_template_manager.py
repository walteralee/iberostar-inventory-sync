"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    excel_template_manager.py

Descripción:
    Servicio encargado de preparar automáticamente los Excel
    mensuales a partir de las plantillas.
"""

from pathlib import Path
import shutil

from config.constants import (
    EXCEL_TEMPLATES,
    MONTHS,
    EXCEL_EXTENSION,
)
from config.settings import (
    INPUT_EXCELS_DIR,
    TEMPLATES_DIR,
)


class ExcelTemplateManager:
    """
    Gestiona la creación automática de los Excel mensuales.
    """

    def ensure_month(
        self,
        year: int,
        month: int,
    ) -> Path:
        """
        Garantiza que existen los Excel correspondientes al
        mes indicado.

        Args:
            year: Año.
            month: Mes (1-12).

        Returns:
            Carpeta donde se encuentran los Excel del mes.
        """

        month_name = self._get_month_name(month)

        month_directory = INPUT_EXCELS_DIR / str(year) / month_name

        month_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        for template in EXCEL_TEMPLATES:

            source = TEMPLATES_DIR / f"{template}{EXCEL_EXTENSION}"

            destination = (
                month_directory
                / f"{template}_{month_name.title()}_{year}{EXCEL_EXTENSION}"
            )

            if not destination.exists():

                shutil.copy2(
                    source,
                    destination,
                )

        return month_directory

    def get_excel_path(
        self,
        sales_point: str,
        year: int,
        month: int,
    ) -> Path:
        """
        Devuelve la ruta del Excel correspondiente a un punto
        de venta, año y mes.

        Args:
            sales_point: Punto de venta.
            year: Año.
            month: Mes (1-12).

        Returns:
            Ruta completa del archivo Excel.
        """

        month_name = self._get_month_name(month)

        return (
            INPUT_EXCELS_DIR
            / str(year)
            / month_name
            / f"{sales_point}_{month_name.title()}_{year}{EXCEL_EXTENSION}"
        )

    def _get_month_name(
        self,
        month: int,
    ) -> str:
        """
        Devuelve el nombre del mes.
        """

        return MONTHS[month - 1]
