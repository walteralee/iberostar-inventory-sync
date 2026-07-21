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
    EXCEL_EXTENSION,
    EXCEL_TEMPLATES,
    MONTHS,
)
from config.settings import (
    MONTHLY_EXCELS_DIR,
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
            year: Año de los Excel mensuales.
            month: Mes comprendido entre 1 y 12.

        Returns:
            Carpeta donde se encuentran los Excel del mes.

        Raises:
            ValueError: Si el año o el mes no son válidos.
            FileNotFoundError: Si falta alguna plantilla.
        """

        self._validate_year(
            year,
        )

        month_name = self._get_month_name(
            month,
        )

        self._validate_templates()

        month_directory = MONTHLY_EXCELS_DIR / str(year) / month_name

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

        created_count = 0
        existing_count = 0

        for index, template in enumerate(
            EXCEL_TEMPLATES,
            start=1,
        ):

            source = self.get_template_path(
                sales_point=template,
            )

            destination = month_directory / (
                f"{template}_" f"{month_name.title()}_" f"{year}" f"{EXCEL_EXTENSION}"
            )

            if destination.exists():

                existing_count += 1

                print(f"{index:03d} | " f"{destination.name:<50} | " "YA EXISTÍA")

                continue

            shutil.copy2(
                source,
                destination,
            )

            created_count += 1

            print(
                f"{index:03d} | " f"{destination.name:<50} | " "CREADO DESDE PLANTILLA"
            )

        print("-" * 100)
        print(f"Excel comprobados: {len(EXCEL_TEMPLATES)}")
        print(f"Excel creados    : {created_count}")
        print(f"Excel existentes : {existing_count}")
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
        Devuelve la ruta del Excel mensual correspondiente a un
        punto de venta, año y mes.

        Args:
            sales_point: Nombre del punto de venta.
            year: Año del Excel mensual.
            month: Mes comprendido entre 1 y 12.

        Returns:
            Ruta completa del archivo Excel mensual.

        Raises:
            ValueError: Si algún argumento no es válido.
        """

        self._validate_sales_point(
            sales_point,
        )

        self._validate_year(
            year,
        )

        month_name = self._get_month_name(
            month,
        )

        return (
            MONTHLY_EXCELS_DIR
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

        Raises:
            ValueError: Si el punto de venta no es válido.
        """

        self._validate_sales_point(
            sales_point,
        )

        return TEMPLATES_DIR / f"{sales_point}{EXCEL_EXTENSION}"

    def _validate_templates(
        self,
    ) -> None:
        """
        Comprueba que existen todas las plantillas necesarias.

        Raises:
            FileNotFoundError: Si falta una o varias plantillas.
        """

        missing_templates: list[Path] = []

        for template in EXCEL_TEMPLATES:

            template_path = TEMPLATES_DIR / f"{template}{EXCEL_EXTENSION}"

            if not template_path.is_file():

                missing_templates.append(
                    template_path,
                )

        if not missing_templates:

            return

        missing_names = "\n".join(
            f"- {template_path}" for template_path in missing_templates
        )

        raise FileNotFoundError(
            "No se pueden preparar los Excel mensuales porque "
            "faltan las siguientes plantillas:\n"
            f"{missing_names}"
        )

    def _validate_sales_point(
        self,
        sales_point: str,
    ) -> None:
        """
        Comprueba que el punto de venta es válido.

        Args:
            sales_point: Nombre del punto de venta.

        Raises:
            ValueError: Si el punto de venta está vacío o no está
                definido en EXCEL_TEMPLATES.
        """

        if not isinstance(
            sales_point,
            str,
        ):

            raise ValueError("El punto de venta debe ser una cadena de texto.")

        normalized_sales_point = sales_point.strip()

        if not normalized_sales_point:

            raise ValueError("El punto de venta no puede estar vacío.")

        if normalized_sales_point not in EXCEL_TEMPLATES:

            raise ValueError(
                f"Punto de venta no reconocido: " f"{normalized_sales_point}"
            )

    def _validate_year(
        self,
        year: int,
    ) -> None:
        """
        Comprueba que el año es válido.

        Args:
            year: Año que se debe validar.

        Raises:
            ValueError: Si el año no es un número entero válido.
        """

        if not isinstance(
            year,
            int,
        ):

            raise ValueError("El año debe ser un número entero.")

        if year < 2000 or year > 2100:

            raise ValueError(f"Año fuera del rango permitido: {year}")

    def _get_month_name(
        self,
        month: int,
    ) -> str:
        """
        Devuelve el nombre del mes correspondiente.

        Args:
            month: Mes comprendido entre 1 y 12.

        Returns:
            Nombre del mes.

        Raises:
            ValueError: Si el mes no es válido.
        """

        if not isinstance(
            month,
            int,
        ):

            raise ValueError("El mes debe ser un número entero.")

        if month < 1 or month > 12:

            raise ValueError(f"Mes fuera del rango permitido: {month}")

        if len(MONTHS) < 12:

            raise ValueError("La constante MONTHS debe contener los 12 meses.")

        return MONTHS[month - 1]
