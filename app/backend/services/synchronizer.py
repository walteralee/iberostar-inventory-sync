"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    synchronizer.py

Descripción:
    Servicio encargado de sincronizar las entregas importadas
    con los Excel mensuales y sus plantillas correspondientes.
"""

from config.constants import DAY_HEADER_ROW

from models.delivery import Delivery

from excel.finder import ExcelFinder
from excel.product_manager import ProductManager
from excel.reader import ExcelReader
from excel.writer import ExcelWriter

from services.excel_template_manager import ExcelTemplateManager
from services.registry import Registry


class Synchronizer:
    """
    Servicio principal de sincronización.
    """

    def __init__(
        self,
        registry: Registry,
    ) -> None:
        """
        Inicializa los servicios necesarios para ejecutar
        la sincronización.

        Args:
            registry: Registro compartido de entregas.
        """

        self.registry = registry

        self.excel_reader = ExcelReader()
        self.excel_finder = ExcelFinder()
        self.excel_writer = ExcelWriter()
        self.template_manager = ExcelTemplateManager()
        self.product_manager = ProductManager()

    def run(
        self,
        deliveries: list[Delivery],
    ) -> None:
        """
        Ejecuta el proceso completo de sincronización.

        Args:
            deliveries: Entregas que deben sincronizarse.
        """

        # ==================================================
        # VALIDAR ENTREGAS RECIBIDAS
        # ==================================================

        if not deliveries:

            print()
            print("=" * 100)
            print("6. INICIO DEL PROCESO DE SINCRONIZACIÓN")
            print("=" * 100)
            print("Entregas recibidas: 0")
            print("Estado            : NO HAY ENTREGAS PENDIENTES")
            print("=" * 100)

            return

        # ==================================================
        # CONTADORES GENERALES
        # ==================================================

        synchronized_delivery_count = 0
        error_delivery_count = 0

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
        print(f"Entregas recibidas: {len(deliveries)}")
        print("-" * 100)

        for index, delivery in enumerate(
            deliveries,
            start=1,
        ):

            sales_point_name = getattr(
                getattr(
                    delivery,
                    "sales_point",
                    None,
                ),
                "name",
                "DESCONOCIDO",
            )

            delivery_date = getattr(
                delivery,
                "delivery_date",
                None,
            )

            formatted_date = (
                delivery_date.strftime("%d/%m/%Y")
                if delivery_date is not None
                else "FECHA DESCONOCIDA"
            )

            print(f"{index:03d} | " f"{formatted_date} | " f"{sales_point_name}")

        print("=" * 100)

        # ==================================================
        # PROCESAMIENTO DE LAS ENTREGAS
        # ==================================================

        for delivery_index, delivery in enumerate(
            deliveries,
            start=1,
        ):

            workbook = None
            template_workbook = None

            sales_point_name = getattr(
                getattr(
                    delivery,
                    "sales_point",
                    None,
                ),
                "name",
                "",
            )

            try:

                # ==========================================
                # VALIDAR ENTREGA
                # ==========================================

                if self.registry.is_synchronized(
                    delivery,
                ):

                    print()
                    print("-" * 100)
                    print("ENTREGA OMITIDA")
                    print("-" * 100)
                    print(
                        f"Fecha          : "
                        f"{delivery.delivery_date.strftime('%d/%m/%Y')}"
                    )
                    print(f"Punto de venta : {sales_point_name}")
                    print("Estado         : YA SINCRONIZADA")
                    print("-" * 100)

                    continue

                if not sales_point_name:

                    raise ValueError("La entrega no contiene un punto de venta válido.")

                if delivery.delivery_date is None:

                    raise ValueError("La entrega no contiene una fecha válida.")

                if not delivery.products:

                    raise ValueError("La entrega no contiene ningún producto.")

                if not self.registry.exists(
                    delivery,
                ):

                    self.registry.register(
                        delivery,
                    )

                    self.registry.save()

                print()
                print()
                print("=" * 100)
                print(
                    f"7. SINCRONIZACIÓN DE LA ENTREGA "
                    f"{delivery_index:03d} DE {len(deliveries):03d}"
                )
                print("=" * 100)
                print(
                    f"Fecha          : "
                    f"{delivery.delivery_date.strftime('%d/%m/%Y')}"
                )
                print(f"Punto de venta : {sales_point_name}")
                print(f"Productos      : {len(delivery.products)}")
                print("=" * 100)

                # ==========================================
                # PREPARAR EXCEL MENSUALES
                # ==========================================

                print()
                print("-" * 100)
                print("7.1 PREPARACIÓN DE LOS EXCEL MENSUALES")
                print("-" * 100)
                print(
                    f"Periodo        : "
                    f"{delivery.delivery_date.month:02d}/"
                    f"{delivery.delivery_date.year}"
                )
                print("Proceso        : Comprobando Excel mensuales...")

                month_directory = self.template_manager.ensure_month(
                    year=delivery.delivery_date.year,
                    month=delivery.delivery_date.month,
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
                    sales_point=sales_point_name,
                    year=delivery.delivery_date.year,
                    month=delivery.delivery_date.month,
                )

                print()
                print("-" * 100)
                print("7.2 APERTURA DEL EXCEL MENSUAL")
                print("-" * 100)
                print(f"Punto de venta : {sales_point_name}")
                print(f"Archivo        : {excel_path.name}")
                print(f"Ruta           : {excel_path}")
                print("Proceso        : Abriendo libro de Excel...")

                workbook, worksheet = self.excel_reader.read(
                    workbook_path=excel_path,
                )

                print("Libro          : ABIERTO CORRECTAMENTE")
                print(f"Hoja utilizada : {worksheet.title}")
                print("-" * 100)

                # ==========================================
                # CONSTRUIR ÍNDICE DE PRODUCTOS
                # ==========================================

                print()
                print("-" * 100)
                print("7.3 PREPARACIÓN DE LA BÚSQUEDA DE PRODUCTOS")
                print("-" * 100)
                print("Proceso        : Construyendo índice código → fila...")

                product_index = self.excel_finder.build_product_index(
                    worksheet=worksheet,
                )

                print(f"Productos      : {len(product_index)}")
                print("Estado         : ÍNDICE DISPONIBLE")
                print("-" * 100)

                # ==========================================
                # LOCALIZAR COLUMNA DEL DÍA
                # ==========================================

                print()
                print("-" * 100)
                print("7.4 LOCALIZACIÓN DEL DÍA DE LA ENTREGA")
                print("-" * 100)

                day = delivery.delivery_date.day

                day_column = self.excel_finder.find_day_column(
                    day=day,
                )

                print(f"Día            : {day}")
                print(f"Columna        : {day_column}")
                print("Proceso        : Validando cabecera del día...")

                day_header_value = worksheet.cell(
                    row=DAY_HEADER_ROW,
                    column=day_column,
                ).value
                if str(day_header_value).strip() != str(day):

                    raise ValueError(
                        f"La columna {day_column} no corresponde "
                        f"al día {day}. "
                        f"Valor encontrado: {day_header_value}"
                    )

                print(f"Cabecera       : {day_header_value}")
                print("Estado         : COLUMNA DEL DÍA VALIDADA")
                print("-" * 100)

                # ==========================================
                # CONTADORES DE LA ENTREGA
                # ==========================================

                written = 0
                created_in_month = 0
                existing_in_month = 0
                created_in_template = 0

                new_products = []

                # ==========================================
                # SINCRONIZAR PRODUCTOS
                # ==========================================

                print()
                print("=" * 100)
                print("7.5 SINCRONIZACIÓN DE PRODUCTOS")
                print("=" * 100)
                print(f"Excel          : {excel_path.name}")
                print(f"Día            : {day}")
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
                    print(f"Cantidad       : {product.quantity}")
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
                # ACTUALIZAR PLANTILLA ORIGINAL
                # ==========================================

                print()
                print("=" * 100)
                print("7.6 ACTUALIZACIÓN DE LA PLANTILLA ORIGINAL")
                print("=" * 100)

                if new_products:

                    template_path = self.template_manager.get_template_path(
                        sales_point=sales_point_name,
                    )

                    print(f"Plantilla      : {template_path.name}")
                    print(f"Ruta           : {template_path}")
                    print(f"Productos nuevos detectados: " f"{len(new_products)}")
                    print("-" * 100)
                    print("Proceso        : Abriendo plantilla original...")

                    template_workbook, template_worksheet = self.excel_reader.read(
                        workbook_path=template_path,
                    )

                    print("Plantilla      : ABIERTA CORRECTAMENTE")
                    print(f"Hoja utilizada : {template_worksheet.title}")

                    print("Proceso        : Construyendo índice " "de la plantilla...")

                    template_product_index = self.excel_finder.build_product_index(
                        worksheet=template_worksheet,
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
                            print("Resultado      : YA EXISTÍA")

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

                # ==========================================
                # GUARDAR EXCEL MENSUAL
                # ==========================================

                print()
                print("=" * 100)
                print("7.7 GUARDADO DEL EXCEL MENSUAL")
                print("=" * 100)

                print(f"Archivo        : {excel_path.name}")
                print(f"Ruta           : {excel_path}")
                print("Proceso        : Guardando cambios...")

                workbook.save(
                    excel_path,
                )

                self.registry.mark_as_synchronized(
                    delivery,
                )

                self.registry.save()

                print("Estado         : GUARDADO CORRECTAMENTE")
                print("=" * 100)

                # ==========================================
                # ACTUALIZAR CONTADORES GENERALES
                # ==========================================

                synchronized_delivery_count += 1
                total_products_written += written
                total_created_in_month += created_in_month
                total_created_in_template += created_in_template

                # ==========================================
                # RESUMEN DE LA ENTREGA
                # ==========================================

                print()
                print("=" * 100)
                print(
                    f"RESUMEN DE LA ENTREGA "
                    f"{delivery_index:03d} DE {len(deliveries):03d}"
                )
                print("=" * 100)
                print(
                    f"Fecha                   : "
                    f"{delivery.delivery_date.strftime('%d/%m/%Y')}"
                )
                print(f"Punto de venta          : {sales_point_name}")
                print(f"Excel mensual           : {excel_path.name}")
                print(f"Productos recibidos     : {len(delivery.products)}")
                print(f"Productos existentes    : {existing_in_month}")
                print(f"Nuevos en Excel mensual : {created_in_month}")
                print(f"Productos escritos      : {written}")
                print(f"Nuevos en plantilla     : {created_in_template}")
                print("Estado                  : SINCRONIZADA CORRECTAMENTE")
                print("=" * 100)

            except Exception as error:

                error_delivery_count += 1

                print()
                print("!" * 100)
                print(
                    f"ERROR DURANTE LA SINCRONIZACIÓN "
                    f"DE LA ENTREGA {delivery_index:03d}"
                )
                print("!" * 100)
                print(f"Punto de venta : {sales_point_name}")
                print("Estado         : ERROR")
                print(f"Tipo           : {type(error).__name__}")
                print(f"Motivo         : {error}")
                print("Resultado      : ENTREGA NO SINCRONIZADA")
                print("!" * 100)

                continue

            finally:

                if template_workbook is not None:

                    template_workbook.close()

                if workbook is not None:

                    workbook.close()

        # ==================================================
        # RESUMEN GENERAL
        # ==================================================

        print()
        print()
        print("=" * 100)
        print("8. RESUMEN GENERAL DE SINCRONIZACIÓN")
        print("=" * 100)
        print(f"Entregas recibidas         : {len(deliveries)}")
        print(f"Entregas sincronizadas     : " f"{synchronized_delivery_count}")
        print(f"Entregas con errores       : " f"{error_delivery_count}")
        print("-" * 100)
        print(f"Productos escritos         : {total_products_written}")
        print(f"Nuevos en Excel mensuales  : {total_created_in_month}")
        print(f"Nuevos en plantillas       : {total_created_in_template}")
        print("-" * 100)

        if error_delivery_count == 0:

            print("Estado final               : SINCRONIZACIÓN COMPLETADA")

        else:

            print("Estado final               : COMPLETADA CON ERRORES")

        print("=" * 100)
