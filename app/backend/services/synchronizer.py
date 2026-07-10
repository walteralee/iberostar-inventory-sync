"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    synchronizer.py

Descripción:
    Servicio principal encargado de sincronizar los albaranes PDF con los
    archivos Excel correspondientes.
"""

from pathlib import Path

from services.excel_template_manager import ExcelTemplateManager

from pdf.parser import PDFParser

from excel.reader import ExcelReader
from excel.finder import ExcelFinder
from excel.writer import ExcelWriter


class Synchronizer:
    """
    Servicio principal de sincronización.
    """

    def __init__(self) -> None:

        self.pdf_parser = PDFParser()

        self.excel_reader = ExcelReader()
        self.excel_finder = ExcelFinder()
        self.excel_writer = ExcelWriter()

        self.template_manager = ExcelTemplateManager()

    def run(
        self,
        pdf_files: list[Path],
    ) -> None:
        """
        Ejecuta el proceso completo de sincronización.

        Args:
            pdf_files: Lista de PDF que deben sincronizarse.
        """

        print(f"PDF encontrados: {len(pdf_files)}")
        print()

        for pdf_file in pdf_files:

            try:

                print("\n" * 2)
                print("=" * 100)
                print("INICIO DE SINCRONIZACIÓN")
                print("=" * 100)
                print(f"PDF: {pdf_file.name}")
                print("=" * 100)

                delivery = self.pdf_parser.parse(pdf_file)

                # ==================================================
                # Ignorar PDFs que no pertenecen a los puntos
                # de venta soportados.
                # ==================================================

                if delivery is None:

                    print("\n" + "!" * 100)
                    print("PDF IGNORADO")
                    print("No pertenece a un punto de venta soportado.")
                    print("!" * 100)

                    continue

                self.template_manager.ensure_month(
                    delivery.delivery_date.year,
                    delivery.delivery_date.month,
                )

                excel_path = self.template_manager.get_excel_path(
                    sales_point=delivery.sales_point.name,
                    year=delivery.delivery_date.year,
                    month=delivery.delivery_date.month,
                )

                workbook, worksheet = self.excel_reader.read(
                    excel_path,
                )

                product_index = self.excel_finder.build_product_index(
                    worksheet,
                )

                day_column = self.excel_finder.find_day_column(
                    delivery.delivery_date.day,
                )

                written = 0
                missing = 0

                for product in delivery.products:

                    row = self.excel_finder.find_product(
                        product_index,
                        product.code,
                    )

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

                workbook.save(
                    excel_path,
                )

                print("\n" + "=" * 100)
                print("RESUMEN DE SINCRONIZACIÓN")
                print("=" * 100)
                print(f"Excel               : {excel_path.name}")
                print(f"Productos en PDF    : {len(delivery.products)}")
                print(f"Productos escritos  : {written}")
                print(f"No encontrados      : {missing}")
                print("=" * 100)

            except Exception as error:

                print("\n" + "!" * 100)
                print(f"ERROR DURANTE LA SINCRONIZACIÓN (PDF {pdf_file.name})")
                print(f"Motivo : {error}")
                print("!" * 100)

                continue

        print("\n" * 2)
        print("=" * 100)
        print("SINCRONIZACIÓN FINALIZADA")
        print("=" * 100)
