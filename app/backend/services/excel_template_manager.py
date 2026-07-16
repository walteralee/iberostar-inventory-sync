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

        month_name = self._get_month_name(
            month,
        )

        month_directory = INPUT_EXCELS_DIR / str(year) / month_name

        print()
        print("=" * 100)
        print("PREPARACIÓN DE LOS EXCEL MENSUALES")
        print("=" * 100)
        print(f"Año              : {year}")
        print(f"Mes              : {month_name.title()}")
        print(f"Carpeta mensual  : {month_directory}")
        print("-" * 100)
        print("Proceso          : Preparando carpeta mensual...")

        month_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        print("Carpeta mensual  : PREPARADA CORRECTAMENTE")
        print("-" * 100)
        print("COMPROBACIÓN DE LOS EXCEL")
        print("-" * 100)

        for index, template in enumerate(
            EXCEL_TEMPLATES,
            start=1,
        ):

            source = TEMPLATES_DIR / f"{template}{EXCEL_EXTENSION}"

            destination = month_directory / (
                f"{template}_" f"{month_name.title()}_" f"{year}" f"{EXCEL_EXTENSION}"
            )

            if not destination.exists():

                shutil.copy2(
                    source,
                    destination,
                )

                print(
                    f"{index:03d} | "
                    f"{destination.name:<50} | "
                    f"CREADO DESDE PLANTILLA"
                )

            else:

                print(f"{index:03d} | " f"{destination.name:<50} | " f"YA EXISTÍA")

        print("-" * 100)
        print(f"Excel comprobados: {len(EXCEL_TEMPLATES)}")
        print(f"Carpeta preparada: {month_directory}")
        print("Estado           : EXCEL MENSUALES DISPONIBLES")
        print("=" * 100)

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

        month_name = self._get_month_name(
            month,
        )

        return (
            INPUT_EXCELS_DIR
            / str(year)
            / month_name
            / (
                f"{sales_point}_"
                f"{month_name.title()}_"
                f"{year}"
                f"{EXCEL_EXTENSION}"
            )
        )

    def get_template_path(
        self,
        sales_point: str,
    ) -> Path:
        """
        Devuelve la ruta de la plantilla original correspondiente
        a un punto de venta.

        Args:
            sales_point: Nombre del punto de venta.

        Returns:
            Ruta completa de la plantilla Excel.
        """

        return TEMPLATES_DIR / f"{sales_point}{EXCEL_EXTENSION}"

    def _get_month_name(
        self,
        month: int,
    ) -> str:
        """
        Devuelve el nombre del mes.
        """

        return MONTHS[month - 1]
