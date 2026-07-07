"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    parser.py

Descripción:
    Convierte un PDF de Iberostar en un objeto Delivery.
"""

from datetime import datetime
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

    PRODUCT_FORMATS = {
        "BOTELLA",
        "LITRO",
        "UNIDAD",
        "PAQUETE",
        "BRIK",
        "KILO",
        "MANOJO",
        "CAJA",
        "BOLSA",
        "BANDEJA",
        "SACO",
        "BOTE",
        "LATA",
        "TARRINA",
    }

    def parse(self, pdf_path: Path) -> Delivery | None:
        """
        Convierte un PDF en un objeto Delivery.

        Si el PDF no pertenece a ninguno de los puntos de venta del
        sistema, devuelve None para que el Synchronizer lo ignore.
        """

        text = self._extract_text(pdf_path)

        lines = self._merge_broken_lines(text)

        lines = self._clean_product_lines(lines)

        lines = self._remove_invalid_lines(lines)

        delivery_date = self._extract_date(lines)

        sales_point = self._extract_sales_point(lines)

        if sales_point is None:
            return None

        products = self._extract_products(lines)

        return Delivery(
            sales_point=sales_point,
            delivery_date=delivery_date,
            products=products,
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
        Une líneas partidas del PDF.
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

        for line in text.splitlines():

            line = line.strip()

            if not line:
                continue

            # --------------------------------------------------
            # Fin de la tabla de productos.
            # --------------------------------------------------

            if line.startswith("Total"):

                if current:
                    merged.append(current)
                    current = ""

                merged.append(line)

                break

            # --------------------------------------------------
            # Ignorar pies de página y cabeceras repetidas.
            # --------------------------------------------------

            if line.startswith(ignore_prefixes):
                continue

            # Ignorar "01/07/2026 13:29 1/2"
            if re.match(r"^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s+\d+/\d+$", line):
                continue

            # --------------------------------------------------
            # Nuevo producto.
            # --------------------------------------------------

            if re.match(r"^\d+\s", line):

                if current:
                    merged.append(current)

                current = line

            # --------------------------------------------------
            # Continuación del producto anterior.
            # --------------------------------------------------

            else:

                if current:
                    current += " " + line
                else:
                    merged.append(line)

        if current:
            merged.append(current)

        return merged

    def _clean_product_lines(
        self,
        lines: list[str],
    ) -> list[str]:
        """
        Limpia las líneas de productos eliminando cualquier texto
        que aparezca después de las dos últimas columnas numéricas.

        Ejemplo:

        27935 ... 0,691 16,590 POMELO S/R 20 CL

        →

        27935 ... 0,691 16,590
        """

        cleaned = []

        pattern = re.compile(r"^(.*?\d+,\d+\s+\d+,\d+)")

        for line in lines:

            if not re.match(r"^\d+\s", line):
                cleaned.append(line)
                continue

            match = pattern.match(line)

            if match:
                cleaned.append(match.group(1))
            else:
                cleaned.append(line)

        return cleaned

    def _remove_invalid_lines(
        self,
        lines: list[str],
    ) -> list[str]:
        """
        Elimina cualquier línea de producto que no termine
        exactamente en dos números decimales.

        Una línea válida debe terminar en:

            0,691 16,590
        """

        cleaned = []

        product_pattern = re.compile(r"\d+,\d+\s+\d+,\d+$")

        for line in lines:

            # ----------------------------------------------
            # Las líneas que NO son productos se conservan.
            # ----------------------------------------------

            if not re.match(r"^\d+", line):
                cleaned.append(line)
                continue

            # ----------------------------------------------
            # Solo conservamos productos completos.
            # ----------------------------------------------

            if product_pattern.search(line):
                cleaned.append(line)

        print("\n" + "=" * 100)
        print("LÍNEAS TRAS ELIMINAR PRODUCTOS INVÁLIDOS")
        print("=" * 100)

        for index, line in enumerate(cleaned, start=1):
            print(f"{index:03d} | {line}")

        print("=" * 100)
        print(f"TOTAL DE LÍNEAS: {len(cleaned)}")
        print("=" * 100)

        return cleaned

    def _extract_date(self, lines: list[str]):
        """
        Extrae la fecha del albarán.
        """

        for line in lines:

            if line.startswith("Fecha:"):

                value = line.replace("Fecha:", "").strip()

                return datetime.strptime(value, "%d/%m/%Y").date()

        raise ValueError("Fecha no encontrada.")

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

        decimal_pattern = re.compile(r"^\d+,\d+$")
        integer_pattern = re.compile(r"^\d+$")

        for line_number, line in enumerate(lines, start=1):

            if line.startswith("Total"):
                continue

            tokens = line.split()

            if len(tokens) < 4:
                continue

            if not tokens[0].isdigit():
                continue

            code = int(tokens[0])

            # --------------------------------------------------
            # Localizar Precio y Total (dos últimos decimales)
            # --------------------------------------------------

            decimal_indexes = [
                index
                for index, token in enumerate(tokens)
                if decimal_pattern.match(token)
            ]

            if len(decimal_indexes) < 2:

                print(
                    f"[ERROR] Línea {line_number}: "
                    f"No se encontraron Precio y Total."
                )

                print(f"        {line}")

                continue

            price_index = decimal_indexes[-2]

            # --------------------------------------------------
            # Buscar enteros inmediatamente anteriores
            # --------------------------------------------------

            integers = []

            index = price_index - 1

            while index >= 0 and integer_pattern.match(tokens[index]):
                integers.insert(0, tokens[index])
                index -= 1

            if not integers:

                print(f"[ERROR] Línea {line_number}: " f"No se encontró la cantidad.")

                print(f"        {line}")

                continue

            # Caso normal:
            # Cantidad Cant.Sol Vale Precio Total

            if len(integers) >= 3:
                quantity = float(integers[-3])

            # Caso Star Café:
            # Cantidad Precio Total

            else:
                quantity = float(integers[-1])

            product = Product(
                code=code,
                quantity=quantity,
            )

            products.append(product)

            print(
                f"{len(products):03d} | "
                f"Código: {product.code:<8} | "
                f"Cantidad: {product.quantity}"
            )

        print("=" * 100)
        print(f"TOTAL DE PRODUCTOS: {len(products)}")
        print("=" * 100)

        return products
