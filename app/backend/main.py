"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    main.py

Descripción:
    Punto de entrada de la aplicación.
"""

from services.importer import Importer
from services.registry import Registry
from services.synchronizer import Synchronizer


def main() -> None:
    """
    Ejecuta el proceso completo de importación
    y sincronización.
    """

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

    if not deliveries:

        print()
        print("=" * 100)
        print("PROCESO FINALIZADO")
        print("=" * 100)
        print("No existen entregas pendientes de sincronizar.")
        print("=" * 100)

        return

    # ==================================================
    # SINCRONIZACIÓN
    # ==================================================

    synchronizer.run(
        deliveries,
    )


if __name__ == "__main__":
    main()
