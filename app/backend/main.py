"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    main.py

Descripción:
    Punto de entrada de la aplicación.

    Orquesta la importación y la sincronización, y concentra en un único
    lugar toda la salida por consola del proceso completo: el título
    introductorio y, al finalizar, el resumen de errores/advertencias
    (o el mensaje de éxito si no hubo ninguno).
"""

from services.importer import Importer
from services.registry import Registry
from services.synchronizer import Synchronizer
from utils.delivery_identity import build_delivery_key


def main() -> None:
    """
    Ejecuta el proceso completo de importación
    y sincronización.
    """

    _print_intro()

    # ==================================================
    # SERVICIOS COMPARTIDOS
    # ==================================================

    registry = Registry()

    importer = Importer(
        registry=registry,
    )

    synchronizer = Synchronizer(
        registry=registry,
    )

    # ==================================================
    # IMPORTACIÓN
    # ==================================================

    deliveries = importer.run()
    import_summary = importer.last_summary

    # ==================================================
    # RECUPERACIÓN DE PENDIENTES DE EJECUCIONES ANTERIORES
    # ==================================================
    #
    # Una entrega puede quedar registrada pero sin sincronizar si una
    # ejecución anterior falló o se interrumpió a mitad de camino. Como
    # el Importer solo compara contra lo que trae el Excel de hoy, esa
    # entrega no se retomaría nunca si su Excel de origen no se vuelve
    # a importar. Por eso, aquí se recupera del Registry todo lo que
    # sigue pendiente y se añade a la tanda de esta sincronización.

    known_delivery_keys = {
        build_delivery_key(delivery) for delivery in deliveries
    }

    recovered_deliveries, pending_warnings = registry.get_pending_deliveries()

    recovered_deliveries = [
        delivery
        for delivery in recovered_deliveries
        if build_delivery_key(delivery) not in known_delivery_keys
    ]

    deliveries = deliveries + recovered_deliveries

    # ==================================================
    # SINCRONIZACIÓN
    # ==================================================

    synchronization_totals = synchronizer.run(deliveries)

    # Resumen final combinando importación, Registry y sincronización.
    _print_final_summary(
        import_summary=import_summary,
        pending_warnings=pending_warnings,
        synchronization_totals=synchronization_totals,
    )


def _print_intro() -> None:
    """
    Título introductorio del proceso.
    """

    print("=" * 100)
    print("IBEROSTAR - SINCRONIZADOR DE INVENTARIO")
    print("=" * 100)


def _print_final_summary(
    import_summary,
    pending_warnings,
    synchronization_totals,
) -> None:
    """
    Resumen consolidado de todos los errores y advertencias de la
    importación y la sincronización, o un mensaje de éxito cuando no
    hubo ninguno.
    """

    print("=" * 100)
    print("RESUMEN FINAL")
    print("=" * 100)

    has_incidents = (
        import_summary.has_incidents()
        or bool(pending_warnings)
        or bool(synchronization_totals.error_deliveries)
    )

    if not has_incidents:
        print("Proceso completado sin errores ni advertencias.")
        print("=" * 100)
        return

    print("IMPORTACIÓN")
    print(f"  Excel seleccionados         : {import_summary.selected_file_count}")
    print(f"  Excel procesados            : {import_summary.processed_file_count}")
    print(f"  Excel descartados por error : {import_summary.file_error_count}")
    for message in import_summary.file_error_messages:
        print(f"    - {message}")

    print(f"  Filas leídas                : {import_summary.total_rows_read}")
    print(f"  Filas válidas               : {import_summary.valid_row_count}")
    print(f"  Grupos no admitidos         : {import_summary.ignored_group_count}")
    print(
        "  Puntos de venta ignorados   : "
        f"{import_summary.ignored_sales_point_count}"
    )
    print(
        "  Cantidades a cero           : "
        f"{import_summary.ignored_zero_quantity_count}"
    )

    print(f"  Duplicados entre Excel      : {import_summary.ignored_duplicate_count}")
    for message in import_summary.duplicate_messages:
        print(f"    - {message}")

    print(f"  Filas con error de parseo   : {import_summary.row_error_count}")
    for message in import_summary.row_error_messages:
        print(f"    - {message}")

    if import_summary.discrepancy_messages:
        print(
            "  Discrepancias de datos      : "
            f"{len(import_summary.discrepancy_messages)}"
        )
        for message in import_summary.discrepancy_messages:
            print(f"    - {message}")

    if import_summary.thousands_format_messages:
        print(
            "  Avisos de formato numérico  : "
            f"{len(import_summary.thousands_format_messages)}"
        )
        for message in import_summary.thousands_format_messages:
            print(f"    - {message}")

    print(f"  Conflictos de Registry      : {import_summary.conflict_count}")
    for message in import_summary.conflict_messages:
        print(f"    - {message}")

    print("-" * 100)
    print("REGISTRY")
    print(
        "  Entregas pendientes no reconstruibles: "
        f"{len(pending_warnings)}"
    )
    for message in pending_warnings:
        print(f"    - {message}")

    print("-" * 100)
    print("SINCRONIZACIÓN")
    print(
        "  Entregas con error de sincronización: "
        f"{synchronization_totals.error_deliveries}"
    )
    for message in synchronization_totals.error_messages:
        print(f"    - {message}")

    print("=" * 100)


if __name__ == "__main__":
    main()
