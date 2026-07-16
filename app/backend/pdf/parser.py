"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    parser.py

Descripción:
    Convierte un PDF de Iberostar en un objeto Delivery.
"""

from datetime import date, datetime
from pathlib import Path
import re

import pdfplumber

from config.constants import SALES_POINT_MAPPING
from models.delivery import Delivery
from models.product import Product
from models.sales_point import SalesPoint


class PDFParser:
    """
    Convierte un PDF de Iberostar en un objeto Delivery.
    """

    def parse(
        self,
        pdf_path: Path,
    ) -> Delivery | None:
        """
        Convierte un PDF en un objeto Delivery.

        Si el PDF no pertenece a ninguno de los puntos de venta
        soportados, devuelve None.
        """

        print()
        print("=" * 100)
        print("ANÁLISIS DEL ALBARÁN PDF")
        print("=" * 100)
        print(f"Archivo       : {pdf_path.name}")
        print("Estado        : INICIANDO LECTURA")
        print("-" * 100)

        print("Proceso       : Extrayendo texto del PDF...")

        raw_text = self._extract_text(
            pdf_path,
        )

        print("Texto         : EXTRAÍDO CORRECTAMENTE")
        print("Proceso       : Uniendo líneas fragmentadas...")

        merged_lines = self._merge_broken_lines(
            raw_text,
        )

        print(f"Líneas unidas : {len(merged_lines)}")
        print("Proceso       : Limpiando y normalizando líneas...")

        lines = self._clean_lines(
            merged_lines,
        )

        print(f"Líneas válidas: {len(lines)}")
        print("-" * 100)

        print("Proceso       : Extrayendo código IBS...")

        ibs_code = self._extract_ibs_code(
            lines,
        )

        print(f"Código IBS    : {ibs_code}")
        print("Proceso       : Identificando punto de venta...")

        sales_point = self._extract_sales_point(
            lines,
        )

        if sales_point is None:

            print()
            print("!" * 100)
            print("PDF NO SOPORTADO")
            print("!" * 100)
            print(f"Archivo       : {pdf_path.name}")
            print(f"Código IBS    : {ibs_code}")
            print("Punto de venta: NO SOPORTADO")
            print("Resultado     : PDF IGNORADO")
            print("!" * 100)

            return None

        delivery_date = self._extract_date(
            lines,
        )

        print(f"Punto de venta: {sales_point.name}")
        print(f"Fecha         : {delivery_date.strftime('%d/%m/%Y')}")
        print("Metadatos     : EXTRAÍDOS CORRECTAMENTE")
        print("=" * 100)

        products = self._extract_products(
            lines,
        )

        print()
        print("=" * 100)
        print("RESULTADO DEL ANÁLISIS")
        print("=" * 100)
        print(f"Archivo              : {pdf_path.name}")
        print(f"Código IBS           : {ibs_code}")
        print(f"Fecha                : {delivery_date.strftime('%d/%m/%Y')}")
        print(f"Punto de venta       : {sales_point.name}")
        print(f"Productos encontrados: {len(products)}")
        print("Estado               : PDF PARSEADO CORRECTAMENTE")
        print("=" * 100)

        return Delivery(
            ibs_code=ibs_code,
            sales_point=sales_point,
            delivery_date=delivery_date,
            products=products,
        )

    # ======================================================
    # PRIVATE
    # ======================================================

    def _extract_text(
        self,
        pdf_path: Path,
    ) -> str:
        """
        Extrae el texto completo del PDF.
        """

        pages = []

        with pdfplumber.open(
            pdf_path,
        ) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:
                    pages.append(
                        text,
                    )

        return "\n".join(
            pages,
        )

    def _merge_broken_lines(
        self,
        text: str,
    ) -> list[str]:
        """
        Une las líneas partidas del PDF.
        """

        merged = []
        current = ""

        ignore_prefixes = (
            "Revisado Por:",
            "Entregado Por:",
            "Recibido Por:",
            "Firma:",
            "Nombre:",
            "Cód.Art.",
        )

        page_pattern = re.compile(
            r"^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s+\d+/\d+$"
        )

        product_pattern = re.compile(
            r"^\d+\s"
        )

        for line in text.splitlines():

            line = line.strip()

            if not line:
                continue

            if line.startswith("Total"):

                if current:
                    merged.append(
                        current,
                    )
                    current = ""

                merged.append(
                    line,
                )

                break

            if (
                line.startswith(ignore_prefixes)
                or page_pattern.match(line)
            ):
                continue

            if product_pattern.match(line):

                if current:
                    merged.append(
                        current,
                    )

                current = line

            elif current:

                current += " " + line

            else:

                merged.append(
                    line,
                )

        if current:
            merged.append(
                current,
            )

        return merged

    def _clean_lines(
        self,
        lines: list[str],
    ) -> list[str]:
        """
        Normaliza las líneas de productos.

        - Elimina el texto sobrante tras las dos últimas columnas numéricas.
        - Descarta productos incompletos.
        """

        cleaned = []

        trim_pattern = re.compile(
            r"^(.*?\d+,\d+\s+\d+,\d+)"
        )

        valid_product_pattern = re.compile(
            r"\d+,\d+\s+\d+,\d+$"
        )

        product_pattern = re.compile(
            r"^\d+\s"
        )

        for line in lines:

            if not product_pattern.match(line):
                cleaned.append(
                    line,
                )
                continue

            match = trim_pattern.match(
                line,
            )

            if not match:
                continue

            line = match.group(1)

            if valid_product_pattern.search(line):
                cleaned.append(
                    line,
                )

        return cleaned

    def _extract_date(
        self,
        lines: list[str],
    ) -> date:
        """
        Extrae la fecha del albarán.
        """

        for line in lines:

            if line.startswith("Fecha:"):

                value = line.replace(
                    "Fecha:",
                    "",
                ).strip()

                return datetime.strptime(
                    value,
                    "%d/%m/%Y",
                ).date()

        raise ValueError(
            "Fecha no encontrada.",
        )

    def _extract_ibs_code(
        self,
        lines: list[str],
    ) -> int:
        """
        Extrae el código IBS del albarán.
        """

        for line in lines:

            if line.startswith("Cód Ibs.:"):

                value = line.replace(
                    "Cód Ibs.:",
                    "",
                ).strip()

                return int(
                    value,
                )

        raise ValueError(
            "Código IBS no encontrado.",
        )

    def _extract_sales_point(
        self,
        lines: list[str],
    ) -> SalesPoint | None:
        """
        Extrae el punto de venta.

        Si el departamento no pertenece a los puntos de venta
        soportados, devuelve None.
        """

        for line in lines:

            if line.startswith("Dpto. Destino:"):

                value = line.replace(
                    "Dpto. Destino:",
                    "",
                ).strip()

                if value not in SALES_POINT_MAPPING:
                    return None

                return SalesPoint(
                    SALES_POINT_MAPPING[value],
                )

        raise ValueError(
            "Departamento no encontrado.",
        )

    def _extract_products(
        self,
        lines: list[str],
    ) -> list[Product]:
        """
        Extrae todos los productos del albarán.
        """

        print()
        print("=" * 100)
        print("LISTA DE PRODUCTOS ENCONTRADOS")
        print("=" * 100)
        print(
            f"{'N.º':<5}"
            f"{'CÓDIGO':<10}"
            f"{'NOMBRE':<44}"
            f"{'FORMATO':<12}"
            f"{'PRECIO':>12}"
            f"{'CANTIDAD':>13}"
        )
        print("-" * 100)

        products: list[Product] = []
        product_errors: list[str] = []

        candidate_count = 0

        for line_number, line in enumerate(
            lines,
            start=1,
        ):

            if line.startswith("Total"):
                continue

            tokens = line.split()

            if len(tokens) < 4:
                continue

            if not tokens[0].isdigit():
                continue

            candidate_count += 1

            try:

                (
                    code,
                    name,
                    product_format,
                    price,
                    quantity,
                ) = self._extract_product_data(
                    tokens,
                    line,
                    line_number,
                )

            except ValueError as error:

                product_errors.append(
                    str(error),
                )

                continue

            product = Product(
                code=code,
                name=name,
                format=product_format,
                price=price,
                quantity=quantity,
            )

            products.append(
                product,
            )

            display_name = product.name

            if len(display_name) > 41:
                display_name = display_name[:38] + "..."

            print(
                f"{len(products):03d}  "
                f"{product.code:<10}"
                f"{display_name:<44}"
                f"{product.format:<12}"
                f"{product.price:>12.3f}"
                f"{product.quantity:>13.2f}"
            )

        if not products:

            print("No se encontraron productos válidos.")

        print("-" * 100)
        print(f"Líneas de producto detectadas : {candidate_count}")
        print(f"Productos extraídos           : {len(products)}")
        print(f"Productos con errores         : {len(product_errors)}")
        print("=" * 100)

        if product_errors:

            print()
            print("!" * 100)
            print("ERRORES DURANTE LA EXTRACCIÓN DE PRODUCTOS")
            print("!" * 100)

            for index, error in enumerate(
                product_errors,
                start=1,
            ):

                print(f"ERROR {index:03d}")
                print("-" * 100)
                print(error)
                print("-" * 100)

            print(f"Total de errores: {len(product_errors)}")
            print("!" * 100)

        return products

    def _extract_product_data(
        self,
        tokens: list[str],
        line: str,
        line_number: int,
    ) -> tuple[int, str, str, float, float]:
        """
        Extrae toda la información de un producto.
        """

        decimal_pattern = re.compile(
            r"^\d+,\d+$"
        )

        integer_pattern = re.compile(
            r"^\d+$"
        )

        decimal_indexes = [
            index
            for index, token in enumerate(tokens)
            if decimal_pattern.match(token)
        ]

        if len(decimal_indexes) < 2:
            raise ValueError(
                f"Línea {line_number}: "
                f"No se encontraron el precio y el total.\n"
                f"Contenido: {line}"
            )

        price_index = decimal_indexes[-2]

        index = price_index - 1

        while (
            index >= 0
            and integer_pattern.match(tokens[index])
        ):
            index -= 1

        if index < 1:
            raise ValueError(
                f"Línea {line_number}: "
                f"No se encontró el formato.\n"
                f"Contenido: {line}"
            )

        format_index = index
        product_format = tokens[format_index]

        quantity_index = format_index + 1

        if (
            quantity_index >= len(tokens)
            or not integer_pattern.match(tokens[quantity_index])
        ):
            raise ValueError(
                f"Línea {line_number}: "
                f"No se encontró la cantidad.\n"
                f"Contenido: {line}"
            )

        quantity = float(
            tokens[quantity_index],
        )

        name = " ".join(
            tokens[1:format_index],
        )

        price = float(
            tokens[price_index].replace(
                ",",
                ".",
            )
        )

        code = int(
            tokens[0],
        )

        return (
            code,
            name,
            product_format,
            price,
            quantity,
        )