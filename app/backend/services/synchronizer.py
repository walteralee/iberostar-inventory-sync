"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    synchronizer.py

Descripción:
    Servicio principal encargado de sincronizar los albaranes PDF con los
    archivos Excel correspondientes.
"""

from config.constants import MONTHS
from config.settings import (
    INPUT_PDFS_DIR,
    INPUT_EXCELS_DIR,
)

from pdf.parser import PDFParser

from excel.reader import ExcelReader
from excel.finder import ExcelFinder
from excel.writer import ExcelWriter

from utils.filesystem import get_pdf_files


class Synchronizer:
    """
    Servicio principal de sincronización.
    """

    def __init__(self) -> None:

        self.pdf_parser = PDFParser()

        self.excel_reader = ExcelReader()
        self.excel_finder = ExcelFinder()
        self.excel_writer = ExcelWriter()

    def run(self) -> None:
        """
        Ejecuta el proceso completo de sincronización.
        """

        pdf_files = get_pdf_files(INPUT_PDFS_DIR)

        if not pdf_files:
            print("No se encontraron archivos PDF.")
            return

        print(f"PDF encontrados: {len(pdf_files)}")
        print()

        for pdf_file in pdf_files:

            print("\n" * 2)
            print("=" * 100)
            print("INICIO DE SINCRONIZACIÓN")
            print("=" * 100)
            print(f"PDF: {pdf_file.name}")
            print("=" * 100)

            delivery = self.pdf_parser.parse(pdf_file)

            # ==================================================
            # Ignorar PDFs que no pertenecen a los cinco
            # puntos de venta soportados.
            # ==================================================

            if delivery is None:
                print("\n" + "!" * 100)
                print("PDF IGNORADO")
                print("No pertenece a un punto de venta soportado.")
                print("!" * 100)
                continue

            month = MONTHS[delivery.delivery_date.month - 1]

            year = str(delivery.delivery_date.year)

            excel_name = (
                f"{delivery.sales_point.name}_" f"{month.title()}_" f"{year}.xlsx"
            )

            excel_path = INPUT_EXCELS_DIR / excel_name

            if not excel_path.exists():

                print("\n" + "!" * 100)
                print("EXCEL NO ENCONTRADO")
                print(f"Archivo: {excel_name}")
                print("!" * 100)
                continue

            workbook, worksheet = self.excel_reader.read(excel_path)

            product_index = self.excel_finder.build_product_index(worksheet)

            day_column = self.excel_finder.find_day_column(delivery.delivery_date.day)

            written = 0
            missing = 0

            for product in delivery.products:

                row = self.excel_finder.find_product(product_index, product.code)

                if row is None:

                    print(f"[NO ENCONTRADO] Código: {product.code}")

                    missing += 1

                    continue

                self.excel_writer.write(
                    worksheet=worksheet,
                    row=row,
                    column=day_column,
                    quantity=product.quantity,
                )

                written += 1

            workbook.save(excel_path)

            print("\n" + "=" * 100)
            print("RESUMEN DE SINCRONIZACIÓN")
            print("=" * 100)
            print(f"Excel               : {excel_name}")
            print(f"Productos en PDF    : {len(delivery.products)}")
            print(f"Productos escritos  : {written}")
            print(f"No encontrados      : {missing}")
            print("=" * 100)

        print("\n" * 2)
        print("=" * 100)
        print("SINCRONIZACIÓN FINALIZADA")
        print("=" * 100)
