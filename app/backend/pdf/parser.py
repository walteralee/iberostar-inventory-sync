"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    parser.py

Descripción:
    Convierte un PDF de Iberostar en un objeto Delivery.
"""

from pathlib import Path
import re

import pdfplumber

from config.constants import SALES_POINT_MAPPING
from models.delivery import Delivery
from models.product import Product
from models.sales_point import SalesPoint
from datetime import datetime, date


class PDFParser:
    """
    Convierte un PDF de Iberostar en un objeto Delivery.
    """

    def parse(self, pdf_path: Path) -> Delivery | None:
        """
        Convierte un PDF en un objeto Delivery.

        Si el PDF no pertenece a ninguno de los puntos de venta
        soportados, devuelve None.
        """

        lines = self._clean_lines(
            self._merge_broken_lines(self._extract_text(pdf_path))
        )

        ibs_code = self._extract_ibs_code(lines)

        sales_point = self._extract_sales_point(lines)

        if sales_point is None:
            return None

        return Delivery(
            ibs_code=ibs_code,
            sales_point=sales_point,
            delivery_date=self._extract_date(lines),
            products=self._extract_products(lines),
        )

    # ======================================================
    # PRIVATE
    # ======================================================

    def _extract_text(self, pdf_path: Path) -> str:
        """
        Extrae el texto completo del PDF.
        """

        pages = []

        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:
                    pages.append(text)

        return "\n".join(pages)

    def _merge_broken_lines(self, text: str) -> list[str]:
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

        page_pattern = re.compile(r"^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s+\d+/\d+$")

        product_pattern = re.compile(r"^\d+\s")

        for line in text.splitlines():

            line = line.strip()

            if not line:
                continue

            if line.startswith("Total"):

                if current:
                    merged.append(current)
                    current = ""  # <-- Añade esta línea

                merged.append(line)
                break

            if line.startswith(ignore_prefixes) or page_pattern.match(line):
                continue

            if product_pattern.match(line):

                if current:
                    merged.append(current)

                current = line

            elif current:

                current += " " + line

            else:

                merged.append(line)

        if current:
            merged.append(current)

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

        trim_pattern = re.compile(r"^(.*?\d+,\d+\s+\d+,\d+)")

        valid_product_pattern = re.compile(r"\d+,\d+\s+\d+,\d+$")

        product_pattern = re.compile(r"^\d+\s")

        for line in lines:

            if not product_pattern.match(line):
                cleaned.append(line)
                continue

            match = trim_pattern.match(line)

            if not match:
                continue

            line = match.group(1)

            if valid_product_pattern.search(line):
                cleaned.append(line)

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

                value = line.replace("Fecha:", "").strip()

                return datetime.strptime(value, "%d/%m/%Y").date()

        raise ValueError("Fecha no encontrada.")

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

                return int(value)

        raise ValueError("Código IBS no encontrado.")

    def _extract_sales_point(self, lines: list[str]) -> SalesPoint | None:
        """
        Extrae el punto de venta.

        Si el departamento no pertenece a los cinco puntos de venta
        soportados, devuelve None.
        """

        for line in lines:

            if line.startswith("Dpto. Destino:"):

                value = line.replace("Dpto. Destino:", "").strip()

                if value not in SALES_POINT_MAPPING:
                    return None

                return SalesPoint(SALES_POINT_MAPPING[value])

        raise ValueError("Departamento no encontrado.")

    def _extract_products(self, lines: list[str]) -> list[Product]:
        """
        Extrae todos los productos del albarán.
        """

        print("\n" + "=" * 100)
        print("EXTRACCIÓN DE PRODUCTOS")
        print("=" * 100)

        products = []

        for line_number, line in enumerate(lines, start=1):

            if line.startswith("Total"):
                continue

            tokens = line.split()

            if len(tokens) < 4:
                continue

            if not tokens[0].isdigit():
                continue

            try:

                (
                    code,
                    name,
                    format,
                    price,
                    quantity,
                ) = self._extract_product_data(
                    tokens,
                    line,
                    line_number,
                )

            except ValueError as error:

                print(error)

                continue

            product = Product(
                code=code,
                name=name,
                format=format,
                price=price,
                quantity=quantity,
            )

            products.append(product)

            print(
                f"{len(products):03d} | "
                f"Código: {product.code:<8} | "
                f"Nombre: {product.name:<40} | "
                f"Formato: {product.format:<10} | "
                f"Precio: {product.price:<8} | "
                f"Cantidad: {product.quantity}"
            )

        print("=" * 100)
        print(f"TOTAL DE PRODUCTOS: {len(products)}")
        print("=" * 100)

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

        decimal_pattern = re.compile(r"^\d+,\d+$")
        integer_pattern = re.compile(r"^\d+$")

        decimal_indexes = [
            index for index, token in enumerate(tokens) if decimal_pattern.match(token)
        ]

        if len(decimal_indexes) < 2:
            raise ValueError(
                f"[ERROR] Línea {line_number}: No se encontraron Precio y Total.\n        {line}"
            )

        price_index = decimal_indexes[-2]

        index = price_index - 1

        # Buscar el formato (primera palabra hacia atrás)
        while index >= 0 and integer_pattern.match(tokens[index]):
            index -= 1

        if index < 1:
            raise ValueError(
                f"[ERROR] Línea {line_number}: No se encontró el formato.\n        {line}"
            )

        format_index = index
        product_format = tokens[format_index]

        # Buscar la cantidad (primer número a la derecha del formato)
        quantity_index = format_index + 1

        if (
            quantity_index >= len(tokens)
            or not integer_pattern.match(tokens[quantity_index])
        ):
            raise ValueError(
                f"[ERROR] Línea {line_number}: No se encontró la cantidad.\n        {line}"
            )

        quantity = float(tokens[quantity_index])

        name = " ".join(tokens[1:format_index])

        price = float(tokens[price_index].replace(",", "."))

        code = int(tokens[0])

        return (
            code,
            name,
            product_format,
            price,
            quantity,
        )
