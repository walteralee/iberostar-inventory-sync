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

from excel.finder import ExcelFinder
from excel.product_manager import ProductManager
from excel.reader import ExcelReader
from excel.writer import ExcelWriter
from pdf.parser import PDFParser
from services.excel_template_manager import ExcelTemplateManager


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

        self.product_manager = ProductManager()

    def run(
        self,
        pdf_files: list[Path],
    ) -> None:
        """
        Ejecuta el proceso completo de sincronización.

        Args:
            pdf_files: Lista de PDF que deben sincronizarse.
        """

        # ==================================================
        # CONTADORES GENERALES
        # ==================================================

        synchronized_pdf_count = 0
        ignored_pdf_count = 0
        error_pdf_count = 0

        total_products_written = 0
        total_created_in_month = 0
        total_created_in_template = 0

        # ==================================================
        # INICIO GENERAL
        # ==================================================

        print()
        print("=" * 100)
        print("6. INICIO DEL PROCESO DE SINCRONIZACIÓN")
        print("=" * 100)
        print(f"PDF recibidos desde el importador: {len(pdf_files)}")
        print("-" * 100)

        for index, pdf_file in enumerate(
            pdf_files,
            start=1,
        ):
            print(f"{index:03d} | " f"{pdf_file.name}")

        print("=" * 100)

        # ==================================================
        # PROCESAMIENTO DE LOS PDF
        # ==================================================

        for pdf_index, pdf_file in enumerate(
            pdf_files,
            start=1,
        ):

            try:

                print()
                print()
                print("=" * 100)
                print(
                    f"7. SINCRONIZACIÓN DEL PDF "
                    f"{pdf_index:03d} DE {len(pdf_files):03d}"
                )
                print("=" * 100)
                print(f"Archivo: {pdf_file.name}")
                print("=" * 100)

                # ==========================================
                # PARSEAR PDF
                # ==========================================

                print()
                print("-" * 100)
                print("7.1 LECTURA Y ANÁLISIS DEL PDF")
                print("-" * 100)
                print(f"Archivo        : {pdf_file.name}")
                print("Proceso        : Enviando PDF al parser...")

                delivery = self.pdf_parser.parse(
                    pdf_file,
                )

                # ==========================================
                # PDF NO SOPORTADO
                # ==========================================

                if delivery is None:

                    ignored_pdf_count += 1

                    print()
                    print("!" * 100)
                    print("PDF IGNORADO DURANTE LA SINCRONIZACIÓN")
                    print("!" * 100)
                    print(f"Archivo        : {pdf_file.name}")
                    print("Punto de venta : NO SOPORTADO")
                    print("Resultado      : PDF NO SINCRONIZADO")
                    print("!" * 100)

                    continue

                print()
                print("-" * 100)
                print("DATOS DEL ALBARÁN")
                print("-" * 100)
                print(f"Archivo        : {pdf_file.name}")
                print(f"Código IBS     : {delivery.ibs_code}")
                print(
                    f"Fecha          : "
                    f"{delivery.delivery_date.strftime('%d/%m/%Y')}"
                )
                print(f"Punto de venta : {delivery.sales_point.name}")
                print(f"Productos      : {len(delivery.products)}")
                print("Estado         : PDF ANALIZADO CORRECTAMENTE")
                print("-" * 100)

                # ==========================================
                # PREPARAR EXCEL MENSUALES
                # ==========================================

                print()
                print("-" * 100)
                print("7.2 PREPARACIÓN DE LOS EXCEL MENSUALES")
                print("-" * 100)
                print(
                    f"Periodo        : "
                    f"{delivery.delivery_date.month:02d}/"
                    f"{delivery.delivery_date.year}"
                )
                print("Proceso        : Comprobando Excel mensuales...")

                month_directory = self.template_manager.ensure_month(
                    delivery.delivery_date.year,
                    delivery.delivery_date.month,
                )

                print()
                print("-" * 100)
                print("RESULTADO DE LA PREPARACIÓN MENSUAL")
                print("-" * 100)
                print(f"Carpeta        : {month_directory}")
                print("Estado         : EXCEL MENSUALES DISPONIBLES")
                print("-" * 100)

                # ==========================================
                # LOCALIZAR EXCEL DEL PUNTO DE VENTA
                # ==========================================

                excel_path = self.template_manager.get_excel_path(
                    sales_point=delivery.sales_point.name,
                    year=delivery.delivery_date.year,
                    month=delivery.delivery_date.month,
                )

                print()
                print("-" * 100)
                print("7.3 APERTURA DEL EXCEL MENSUAL")
                print("-" * 100)
                print(f"Punto de venta : {delivery.sales_point.name}")
                print(f"Archivo        : {excel_path.name}")
                print(f"Ruta           : {excel_path}")
                print("Proceso        : Abriendo libro de Excel...")

                workbook, worksheet = self.excel_reader.read(
                    excel_path,
                )

                print("Libro          : ABIERTO CORRECTAMENTE")
                print(f"Hoja utilizada : {worksheet.title}")
                print("-" * 100)

                # ==========================================
                # CONSTRUIR ÍNDICE
                # ==========================================

                print()
                print("-" * 100)
                print("7.4 PREPARACIÓN DE LA BÚSQUEDA DE PRODUCTOS")
                print("-" * 100)
                print("Proceso        : Construyendo índice código → fila...")

                product_index = self.excel_finder.build_product_index(
                    worksheet,
                )

                print(f"Productos      : {len(product_index)}")
                print("Estado         : ÍNDICE DISPONIBLE")
                print("-" * 100)

                # ==========================================
                # BUSCAR COLUMNA DEL DÍA
                # ==========================================

                print()
                print("-" * 100)
                print("7.5 LOCALIZACIÓN DEL DÍA DEL ALBARÁN")
                print("-" * 100)

                day_column = self.excel_finder.find_day_column(
                    delivery.delivery_date.day,
                )

                print(f"Día            : {delivery.delivery_date.day}")
                print(f"Columna        : {day_column}")
                print("Estado         : COLUMNA DEL DÍA LOCALIZADA")
                print("-" * 100)

                # ==========================================
                # CONTADORES DEL PDF
                # ==========================================

                written = 0
                created_in_month = 0
                existing_in_month = 0

                new_products = []

                # ==========================================
                # SINCRONIZAR PRODUCTOS
                # ==========================================

                print()
                print("=" * 100)
                print("7.6 SINCRONIZACIÓN DE PRODUCTOS CON EL EXCEL MENSUAL")
                print("=" * 100)
                print(f"Excel          : {excel_path.name}")
                print(f"Día            : {delivery.delivery_date.day}")
                print(f"Columna        : {day_column}")
                print(f"Productos      : {len(delivery.products)}")
                print("-" * 100)

                for product_index_number, product in enumerate(
                    delivery.products,
                    start=1,
                ):

                    print()
                    print("-" * 100)
                    print(
                        f"PRODUCTO "
                        f"{product_index_number:03d} "
                        f"DE {len(delivery.products):03d}"
                    )
                    print("-" * 100)
                    print(f"Código         : {product.code}")
                    print(f"Nombre         : {product.name}")
                    print(f"Formato        : {product.format}")
                    print(f"Precio         : {product.price}")
                    print(f"Cantidad PDF   : {product.quantity}")
                    print("Proceso        : Buscando producto en el Excel...")

                    row, created = self.product_manager.find_or_create(
                        worksheet=worksheet,
                        product_index=product_index,
                        product=product,
                    )

                    if created:

                        new_products.append(
                            product,
                        )

                        created_in_month += 1

                        print("Existía        : NO")
                        print(f"Fila creada    : {row}")
                        print("Estado         : NUEVO PRODUCTO")

                    else:

                        existing_in_month += 1

                        print("Existía        : SÍ")
                        print(f"Fila utilizada : {row}")
                        print("Estado         : PRODUCTO EXISTENTE")

                    print(
                        f"Escritura      : Acumulando " f"{product.quantity} unidades"
                    )
                    print(f"Destino        : " f"Fila {row} | Columna {day_column}")

                    self.excel_writer.write(
                        worksheet=worksheet,
                        row=row,
                        column=day_column,
                        quantity=product.quantity,
                    )

                    written += 1

                    print("Resultado      : CANTIDAD ESCRITA CORRECTAMENTE")
                    print("-" * 100)

                print()
                print("=" * 100)
                print("RESULTADO DE LA SINCRONIZACIÓN DE PRODUCTOS")
                print("=" * 100)
                print(f"Productos recibidos : {len(delivery.products)}")
                print(f"Productos existentes: {existing_in_month}")
                print(f"Productos nuevos    : {created_in_month}")
                print(f"Productos escritos  : {written}")
                print("=" * 100)

                # ==========================================
                # ACTUALIZAR PLANTILLA
                # ==========================================

                created_in_template = 0

                print()
                print("=" * 100)
                print("7.7 ACTUALIZACIÓN DE LA PLANTILLA ORIGINAL")
                print("=" * 100)

                if new_products:

                    template_path = self.template_manager.get_template_path(
                        sales_point=delivery.sales_point.name,
                    )

                    print(f"Plantilla      : {template_path.name}")
                    print(f"Ruta           : {template_path}")
                    print(f"Productos nuevos detectados: {len(new_products)}")
                    print("-" * 100)
                    print("Proceso        : Abriendo plantilla original...")

                    template_workbook, template_worksheet = self.excel_reader.read(
                        template_path,
                    )

                    print("Plantilla      : ABIERTA CORRECTAMENTE")
                    print(f"Hoja utilizada : {template_worksheet.title}")
                    print("Proceso        : Construyendo índice de la plantilla...")

                    template_product_index = self.excel_finder.build_product_index(
                        template_worksheet,
                    )

                    print(
                        f"Productos indexados en plantilla: "
                        f"{len(template_product_index)}"
                    )

                    for template_product_number, product in enumerate(
                        new_products,
                        start=1,
                    ):

                        print()
                        print("-" * 100)
                        print(
                            f"PRODUCTO NUEVO "
                            f"{template_product_number:03d} "
                            f"DE {len(new_products):03d}"
                        )
                        print("-" * 100)
                        print(f"Código         : {product.code}")
                        print(f"Nombre         : {product.name}")
                        print(
                            "Proceso        : Comprobando producto "
                            "en la plantilla..."
                        )

                        template_row, created = self.product_manager.find_or_create(
                            worksheet=template_worksheet,
                            product_index=template_product_index,
                            product=product,
                        )

                        if created:

                            created_in_template += 1

                            print("Existía        : NO")
                            print(f"Fila creada    : {template_row}")
                            print("Resultado      : AÑADIDO A LA PLANTILLA")

                        else:

                            print("Existía        : SÍ")
                            print(f"Fila utilizada : {template_row}")
                            print("Resultado      : PLANTILLA NO MODIFICADA")

                    print()
                    print("-" * 100)
                    print("RESULTADO DE LA ACTUALIZACIÓN DE LA PLANTILLA")
                    print("-" * 100)
                    print(f"Plantilla          : {template_path.name}")
                    print(f"Productos revisados: {len(new_products)}")
                    print(f"Productos añadidos : {created_in_template}")

                    if created_in_template > 0:

                        print("Proceso            : Guardando plantilla...")

                        template_workbook.save(
                            template_path,
                        )

                        print("Estado             : PLANTILLA GUARDADA")

                    else:

                        print("Estado             : SIN CAMBIOS")
                        print("Guardado           : NO NECESARIO")

                    print("-" * 100)

                else:

                    print(f"Plantilla      : {delivery.sales_point.name}.xlsx")
                    print("Productos nuevos: 0")
                    print("Actualización  : NO NECESARIA")
                    print(
                        "Resultado      : Todos los productos "
                        "ya existían en el Excel mensual."
                    )

                print("=" * 100)

                # ==========================================
                # GUARDAR EXCEL MENSUAL
                # ==========================================

                print()
                print("=" * 100)
                print("7.8 GUARDADO DEL EXCEL MENSUAL")
                print("=" * 100)
                print(f"Archivo        : {excel_path.name}")
                print(f"Ruta           : {excel_path}")
                print("Proceso        : Guardando cambios...")

                workbook.save(
                    excel_path,
                )

                print("Estado         : GUARDADO CORRECTAMENTE")
                print("=" * 100)

                # ==========================================
                # ACTUALIZAR CONTADORES GENERALES
                # ==========================================

                synchronized_pdf_count += 1
                total_products_written += written
                total_created_in_month += created_in_month
                total_created_in_template += created_in_template

                # ==========================================
                # RESUMEN DEL PDF
                # ==========================================

                print()
                print("=" * 100)
                print(f"RESUMEN DEL PDF " f"{pdf_index:03d} DE {len(pdf_files):03d}")
                print("=" * 100)
                print(f"PDF                     : {pdf_file.name}")
                print(f"Código IBS              : {delivery.ibs_code}")
                print(
                    f"Fecha                   : "
                    f"{delivery.delivery_date.strftime('%d/%m/%Y')}"
                )
                print(f"Punto de venta          : {delivery.sales_point.name}")
                print(f"Excel mensual           : {excel_path.name}")
                print(f"Productos en PDF        : {len(delivery.products)}")
                print(f"Productos existentes    : {existing_in_month}")
                print(f"Nuevos en Excel mensual : {created_in_month}")
                print(f"Productos escritos      : {written}")
                print(f"Nuevos en plantilla     : {created_in_template}")
                print("Estado                  : SINCRONIZADO CORRECTAMENTE")
                print("=" * 100)

            except Exception as error:

                error_pdf_count += 1

                print()
                print("!" * 100)
                print(f"ERROR DURANTE LA SINCRONIZACIÓN " f"DEL PDF {pdf_index:03d}")
                print("!" * 100)
                print(f"Archivo        : {pdf_file.name}")
                print("Estado         : ERROR")
                print(f"Tipo           : {type(error).__name__}")
                print(f"Motivo         : {error}")
                print("Resultado      : PDF NO SINCRONIZADO")
                print("!" * 100)

                continue

        # ==================================================
        # RESUMEN GENERAL
        # ==================================================

        print()
        print()
        print("=" * 100)
        print("8. RESUMEN GENERAL DE SINCRONIZACIÓN")
        print("=" * 100)
        print(f"PDF recibidos             : {len(pdf_files)}")
        print(f"PDF sincronizados         : {synchronized_pdf_count}")
        print(f"PDF ignorados             : {ignored_pdf_count}")
        print(f"PDF con errores           : {error_pdf_count}")
        print("-" * 100)
        print(f"Productos escritos        : {total_products_written}")
        print(f"Nuevos en Excel mensuales : {total_created_in_month}")
        print(f"Nuevos en plantillas      : {total_created_in_template}")
        print("-" * 100)

        if error_pdf_count == 0:

            print("Estado final              : SINCRONIZACIÓN COMPLETADA")

        else:

            print("Estado final              : COMPLETADA CON ERRORES")

        print("=" * 100)
