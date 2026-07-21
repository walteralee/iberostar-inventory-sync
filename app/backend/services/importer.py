"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    importer.py

Descripción:
    Servicio encargado de importar los movimientos contenidos
    en los archivos Excel de Economato.

    Las filas válidas se agrupan por fecha y punto de venta.
    Cada grupo se transforma en una entrega con sus productos.
"""

from datetime import date, datetime
from pathlib import Path
from tkinter import Tk, filedialog
import re
import unicodedata

from openpyxl.utils.datetime import from_excel

from config.constants import (
    SALES_POINT_MAPPING,
    SOURCE_DATE_COLUMN,
    SOURCE_EXCEL_EXTENSION,
    SOURCE_FORMAT_COLUMN,
    SOURCE_GROUP_COLUMN,
    SOURCE_HEADER_ROW,
    SOURCE_PRICE_COLUMN,
    SOURCE_PRODUCT_CODE_COLUMN,
    SOURCE_PRODUCT_NAME_COLUMN,
    SOURCE_QUANTITY_COLUMN,
    SOURCE_SALES_POINT_COLUMN,
    SOURCE_SALES_POINT_PREFIX,
    VALID_PRODUCT_GROUPS,
)
from excel.source_reader import SourceReader
from models.delivery import Delivery
from models.product import Product
from models.sales_point import SalesPoint
from services.registry import Registry


class Importer:
    """
    Importa movimientos desde los Excel de Economato y los
    convierte en entregas agrupadas por fecha y punto de venta.
    """

    def __init__(
        self,
        registry: Registry,
    ) -> None:
        """
        Inicializa los servicios necesarios para la importación.

        Args:
            registry: Registro compartido de entregas.
        """

        self.registry = registry
        self.source_reader = SourceReader()

        self._sales_point_mapping = {
            self._normalize_lookup_text(source_name): target_name
            for source_name, target_name in SALES_POINT_MAPPING.items()
        }

        self._valid_product_groups = {
            self._normalize_lookup_text(group) for group in VALID_PRODUCT_GROUPS
        }

        self._normalized_sales_point_prefix = self._normalize_lookup_text(
            SOURCE_SALES_POINT_PREFIX,
        )

    def run(self) -> list[Delivery]:
        """
        Ejecuta el proceso completo de importación.

        Returns:
            Entregas nuevas o pendientes de sincronización.
        """

        # ==================================================
        # SELECCIÓN DE ARCHIVOS
        # ==================================================

        print()
        print("=" * 100)
        print("1. SELECCIÓN DE ARCHIVOS EXCEL")
        print("=" * 100)
        print("Abriendo el explorador de archivos...")
        print("=" * 100)

        excel_files = self._select_excel_files()

        if not excel_files:

            print()
            print("!" * 100)
            print("SELECCIÓN CANCELADA")
            print("!" * 100)
            print("No se seleccionó ningún archivo Excel.")
            print("No hay archivos para importar.")
            print("!" * 100)

            return []

        # ==================================================
        # LISTA DE ARCHIVOS
        # ==================================================

        print()
        print("=" * 100)
        print("2. LISTA DE EXCEL ENCONTRADOS")
        print("=" * 100)

        for index, excel_file in enumerate(
            excel_files,
            start=1,
        ):
            print(f"{index:03d} | {excel_file.name}")

        print("-" * 100)
        print(f"Total de Excel seleccionados: {len(excel_files)}")
        print("=" * 100)

        # La agrupación se realiza de forma global para que dos archivos
        # seleccionados puedan aportar movimientos a la misma entrega.
        grouped_products: dict[
            tuple[date, str],
            dict[str, Product],
        ] = {}

        processed_file_count = 0
        file_error_count = 0

        total_rows_read = 0
        valid_row_count = 0
        ignored_group_count = 0
        ignored_sales_point_count = 0
        ignored_zero_quantity_count = 0
        row_error_count = 0

        # ==================================================
        # PROCESAMIENTO DE CADA EXCEL
        # ==================================================

        print()
        print("=" * 100)
        print("3. PROCESAMIENTO DE LOS EXCEL")
        print("=" * 100)

        for file_index, excel_file in enumerate(
            excel_files,
            start=1,
        ):

            worksheet = None

            try:

                print()
                print("-" * 100)
                print(
                    f"EXCEL {file_index:03d} DE {len(excel_files):03d}"
                    f" | {excel_file.name}"
                )
                print("-" * 100)

                worksheet = self.source_reader.read(
                    excel_file,
                )

                self._validate_worksheet(
                    worksheet.max_column,
                )

                print(f"Hoja abierta    : {worksheet.title}")
                print(f"Filas detectadas: {worksheet.max_row}")
                print(f"Columnas        : {worksheet.max_column}")
                print("Proceso          : Leyendo movimientos...")
                print("-" * 100)

                file_rows_read = 0
                file_valid_rows = 0
                file_row_errors = 0

                for row_number in range(
                    SOURCE_HEADER_ROW + 1,
                    worksheet.max_row + 1,
                ):

                    row_values = self._read_source_row(
                        worksheet=worksheet,
                        row_number=row_number,
                    )

                    if self._is_empty_row(
                        row_values,
                    ):
                        continue

                    file_rows_read += 1
                    total_rows_read += 1

                    try:

                        product_group = self._normalize_lookup_text(
                            row_values["group"],
                        )

                        if product_group not in self._valid_product_groups:

                            ignored_group_count += 1
                            continue

                        sales_point = self._parse_sales_point(
                            row_values["sales_point"],
                        )

                        if sales_point is None:

                            ignored_sales_point_count += 1
                            continue

                        delivery_date = self._parse_date(
                            row_values["date"],
                        )

                        product = self._parse_product(
                            code_value=row_values["code"],
                            name_value=row_values["name"],
                            format_value=row_values["format"],
                            price_value=row_values["price"],
                            quantity_value=row_values["quantity"],
                        )

                        if product.quantity == 0:

                            ignored_zero_quantity_count += 1
                            continue

                        self._add_product(
                            grouped_products=grouped_products,
                            delivery_date=delivery_date,
                            sales_point=sales_point,
                            product=product,
                        )

                        file_valid_rows += 1
                        valid_row_count += 1

                    except (TypeError, ValueError) as error:

                        file_row_errors += 1
                        row_error_count += 1

                        print(
                            f"Fila {row_number:05d} | "
                            f"IGNORADA | {type(error).__name__}: {error}"
                        )

                processed_file_count += 1

                print("-" * 100)
                print(f"Filas leídas     : {file_rows_read}")
                print(f"Filas válidas    : {file_valid_rows}")
                print(f"Errores de fila  : {file_row_errors}")
                print("Estado            : EXCEL PROCESADO")
                print("-" * 100)

            except Exception as error:

                file_error_count += 1

                print()
                print("!" * 100)
                print(f"ERROR DURANTE EL PROCESAMIENTO DEL EXCEL {file_index:03d}")
                print("!" * 100)
                print(f"Archivo : {excel_file.name}")
                print(f"Tipo    : {type(error).__name__}")
                print(f"Motivo  : {error}")
                print("!" * 100)

            finally:

                if worksheet is not None:
                    worksheet.parent.close()

        # ==================================================
        # CONSTRUCCIÓN DE ENTREGAS
        # ==================================================

        print()
        print("=" * 100)
        print("4. CONSTRUCCIÓN DE LAS ENTREGAS")
        print("=" * 100)

        deliveries = self._build_deliveries(
            grouped_products,
        )

        print(f"Grupos encontrados : {len(grouped_products)}")
        print(f"Entregas construidas: {len(deliveries)}")
        print("=" * 100)

        # ==================================================
        # COMPROBACIÓN DEL REGISTRY
        # ==================================================

        print()
        print("=" * 100)
        print("5. COMPROBACIÓN DEL REGISTRY")
        print("=" * 100)

        imported_deliveries: list[Delivery] = []

        imported_count = 0
        pending_count = 0
        existing_count = 0

        for delivery_index, delivery in enumerate(
            deliveries,
            start=1,
        ):

            print()
            print("-" * 100)
            print(f"ENTREGA {delivery_index:03d} " f"DE {len(deliveries):03d}")
            print("-" * 100)
            print("Fecha          : " f"{delivery.delivery_date.strftime('%d/%m/%Y')}")
            print(f"Punto de venta : {delivery.sales_point.name}")
            print(f"Productos      : {len(delivery.products)}")

            if self.registry.exists(
                delivery,
            ):

                if self.registry.is_synchronized(
                    delivery,
                ):

                    existing_count += 1

                    print("Registry       : YA REGISTRADA")
                    print("Sincronización : COMPLETADA")
                    print("Resultado      : ENTREGA OMITIDA")
                    print("-" * 100)

                    continue

                pending_count += 1

                imported_deliveries.append(
                    delivery,
                )

                print("Registry       : YA REGISTRADA")
                print("Sincronización : PENDIENTE")
                print("Resultado      : ENTREGA RECUPERADA")
                print("-" * 100)

                continue

            self.registry.register(
                delivery,
            )

            imported_deliveries.append(
                delivery,
            )

            imported_count += 1

        # ==================================================
        # GUARDADO DEL REGISTRY
        # ==================================================

        print()
        print("=" * 100)
        print("6. GUARDADO DEL REGISTRY")
        print("=" * 100)

        self.registry.save()

        print("Estado: GUARDADO CORRECTAMENTE")
        print("=" * 100)

        # ==================================================
        # RESUMEN
        # ==================================================

        print()
        print("=" * 100)
        print("7. RESUMEN DE IMPORTACIÓN")
        print("=" * 100)
        print(f"Excel seleccionados       : {len(excel_files)}")
        print(f"Excel procesados          : {processed_file_count}")
        print(f"Excel con errores         : {file_error_count}")
        print("-" * 100)
        print(f"Filas leídas              : {total_rows_read}")
        print(f"Filas válidas             : {valid_row_count}")
        print(f"Grupos no admitidos       : {ignored_group_count}")
        print(f"Puntos de venta ignorados : {ignored_sales_point_count}")
        print(f"Cantidades a cero         : {ignored_zero_quantity_count}")
        print(f"Filas con errores         : {row_error_count}")
        print("-" * 100)
        print(f"Entregas construidas      : {len(deliveries)}")
        print(f"Entregas nuevas           : {imported_count}")
        print(f"Entregas pendientes       : {pending_count}")
        print(f"Ya sincronizadas          : {existing_count}")
        print(f"Entregas para sincronizar : {len(imported_deliveries)}")
        print("=" * 100)

        return imported_deliveries

    # ======================================================
    # PRIVATE
    # ======================================================

    def _select_excel_files(self) -> list[Path]:
        """
        Abre el explorador para seleccionar los Excel de origen.
        """

        root = Tk()

        try:

            root.withdraw()
            root.attributes(
                "-topmost",
                True,
            )

            files = filedialog.askopenfilenames(
                title="Seleccionar archivos Excel de Economato",
                filetypes=[
                    (
                        "Archivos Excel",
                        f"*{SOURCE_EXCEL_EXTENSION}",
                    ),
                ],
            )

        finally:

            root.destroy()

        return sorted(Path(file) for file in files)

    def _validate_worksheet(
        self,
        max_column: int,
    ) -> None:
        """
        Comprueba que la hoja contiene todas las columnas necesarias.
        """

        required_last_column = max(
            SOURCE_DATE_COLUMN,
            SOURCE_SALES_POINT_COLUMN,
            SOURCE_GROUP_COLUMN,
            SOURCE_PRODUCT_CODE_COLUMN,
            SOURCE_PRODUCT_NAME_COLUMN,
            SOURCE_FORMAT_COLUMN,
            SOURCE_QUANTITY_COLUMN,
            SOURCE_PRICE_COLUMN,
        )

        if max_column < required_last_column:
            raise ValueError(
                "El Excel no contiene todas las columnas necesarias. "
                f"Se requieren al menos {required_last_column} columnas "
                f"y se encontraron {max_column}."
            )

    def _read_source_row(
        self,
        worksheet,
        row_number: int,
    ) -> dict[str, object]:
        """
        Lee exclusivamente las columnas utilizadas por el proyecto.
        """

        return {
            "date": worksheet.cell(
                row=row_number,
                column=SOURCE_DATE_COLUMN,
            ).value,
            "sales_point": worksheet.cell(
                row=row_number,
                column=SOURCE_SALES_POINT_COLUMN,
            ).value,
            "group": worksheet.cell(
                row=row_number,
                column=SOURCE_GROUP_COLUMN,
            ).value,
            "code": worksheet.cell(
                row=row_number,
                column=SOURCE_PRODUCT_CODE_COLUMN,
            ).value,
            "name": worksheet.cell(
                row=row_number,
                column=SOURCE_PRODUCT_NAME_COLUMN,
            ).value,
            "format": worksheet.cell(
                row=row_number,
                column=SOURCE_FORMAT_COLUMN,
            ).value,
            "quantity": worksheet.cell(
                row=row_number,
                column=SOURCE_QUANTITY_COLUMN,
            ).value,
            "price": worksheet.cell(
                row=row_number,
                column=SOURCE_PRICE_COLUMN,
            ).value,
        }

    def _is_empty_row(
        self,
        row_values: dict[str, object],
    ) -> bool:
        """
        Comprueba si todas las celdas relevantes están vacías.
        """

        return all(
            value is None or str(value).strip() == "" for value in row_values.values()
        )

    def _parse_sales_point(
        self,
        value: object,
    ) -> SalesPoint | None:
        """
        Convierte el destino de Economato en un punto de venta interno.
        """

        normalized_value = self._normalize_lookup_text(
            value,
        )

        if normalized_value.startswith(
            self._normalized_sales_point_prefix,
        ):
            normalized_value = normalized_value[
                len(self._normalized_sales_point_prefix) :
            ].strip()

        mapped_name = self._sales_point_mapping.get(
            normalized_value,
        )

        if mapped_name is None:
            return None

        return SalesPoint(
            name=mapped_name,
        )

    def _parse_date(
        self,
        value: object,
    ) -> date:
        """
        Convierte el valor de fecha del Excel en un objeto date.
        """

        if isinstance(value, datetime):
            return value.date()

        if isinstance(value, date):
            return value

        if isinstance(value, (int, float)) and not isinstance(value, bool):

            parsed_value = from_excel(
                value,
            )

            if isinstance(parsed_value, datetime):
                return parsed_value.date()

            if isinstance(parsed_value, date):
                return parsed_value

        normalized_value = self._require_text(
            value,
            "fecha",
        )

        accepted_formats = (
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%d/%m/%y",
            "%Y/%m/%d",
        )

        for date_format in accepted_formats:

            try:
                return datetime.strptime(
                    normalized_value,
                    date_format,
                ).date()

            except ValueError:
                continue

        raise ValueError(f"Formato de fecha no válido: {value}")

    def _parse_product(
        self,
        code_value: object,
        name_value: object,
        format_value: object,
        price_value: object,
        quantity_value: object,
    ) -> Product:
        """
        Construye un producto a partir de las celdas de una fila.
        """

        return Product(
            code=self._parse_product_code(
                code_value,
            ),
            name=self._require_text(
                name_value,
                "nombre del producto",
            ),
            format=self._require_text(
                format_value,
                "formato del producto",
            ),
            price=self._parse_number(
                price_value,
                "precio",
            ),
            quantity=self._parse_number(
                quantity_value,
                "cantidad",
            ),
        )

    def _parse_product_code(
        self,
        value: object,
    ) -> str:
        """
        Normaliza un código de producto sin convertirlo a float.
        """

        if value is None or isinstance(value, bool):
            raise ValueError("El código del producto está vacío.")

        if isinstance(value, int):
            return str(value)

        if isinstance(value, float):

            if value.is_integer():
                return str(int(value))

            return format(value, "g")

        normalized_value = str(value).strip()

        if not normalized_value:
            raise ValueError("El código del producto está vacío.")

        if re.fullmatch(r"\d+\.0+", normalized_value):
            normalized_value = normalized_value.split(".", maxsplit=1)[0]

        return normalized_value

    def _parse_number(
        self,
        value: object,
        field_name: str,
    ) -> float:
        """
        Convierte un valor numérico del Excel a float.
        """

        if isinstance(value, bool) or value is None:
            raise ValueError(f"El campo {field_name} no es numérico.")

        if isinstance(value, (int, float)):
            return float(value)

        normalized_value = str(value).strip()

        if not normalized_value:
            raise ValueError(f"El campo {field_name} está vacío.")

        normalized_value = normalized_value.replace("€", "")
        normalized_value = normalized_value.replace(" ", "")

        if "," in normalized_value and "." in normalized_value:

            if normalized_value.rfind(",") > normalized_value.rfind("."):
                normalized_value = normalized_value.replace(".", "")
                normalized_value = normalized_value.replace(",", ".")

            else:
                normalized_value = normalized_value.replace(",", "")

        elif "," in normalized_value:
            normalized_value = normalized_value.replace(",", ".")

        try:
            return float(
                normalized_value,
            )

        except ValueError as error:
            raise ValueError(
                f"El campo {field_name} no es numérico: {value}"
            ) from error

    def _add_product(
        self,
        grouped_products: dict[tuple[date, str], dict[str, Product]],
        delivery_date: date,
        sales_point: SalesPoint,
        product: Product,
    ) -> None:
        """
        Añade o acumula un producto dentro de su entrega.
        """

        delivery_key = (
            delivery_date,
            sales_point.name,
        )

        products = grouped_products.setdefault(
            delivery_key,
            {},
        )

        existing_product = products.get(
            product.code,
        )

        if existing_product is None:

            products[product.code] = product
            return

        existing_product.quantity += product.quantity

        # Los datos descriptivos más recientes sustituyen a los anteriores.
        existing_product.name = product.name
        existing_product.format = product.format
        existing_product.price = product.price

    def _build_deliveries(
        self,
        grouped_products: dict[tuple[date, str], dict[str, Product]],
    ) -> list[Delivery]:
        """
        Construye las entregas finales y elimina acumulaciones a cero.
        """

        deliveries: list[Delivery] = []

        for (
            delivery_date,
            sales_point_name,
        ), products_by_code in sorted(
            grouped_products.items(),
            key=lambda item: (
                item[0][0],
                item[0][1],
            ),
        ):

            products = sorted(
                (
                    product
                    for product in products_by_code.values()
                    if product.quantity != 0
                ),
                key=lambda product: product.code,
            )

            if not products:
                continue

            deliveries.append(
                Delivery(
                    sales_point=SalesPoint(
                        name=sales_point_name,
                    ),
                    delivery_date=delivery_date,
                    products=products,
                )
            )

        return deliveries

    def _require_text(
        self,
        value: object,
        field_name: str,
    ) -> str:
        """
        Devuelve un texto limpio y valida que no esté vacío.
        """

        if value is None:
            raise ValueError(f"El campo {field_name} está vacío.")

        normalized_value = " ".join(str(value).strip().split())

        if not normalized_value:
            raise ValueError(f"El campo {field_name} está vacío.")

        return normalized_value

    def _normalize_lookup_text(
        self,
        value: object,
    ) -> str:
        """
        Normaliza textos para comparaciones tolerantes a mayúsculas,
        espacios y acentos.
        """

        if value is None:
            return ""

        normalized_value = " ".join(str(value).strip().upper().split())

        normalized_value = unicodedata.normalize(
            "NFD",
            normalized_value,
        )

        return "".join(
            character
            for character in normalized_value
            if unicodedata.category(character) != "Mn"
        )
