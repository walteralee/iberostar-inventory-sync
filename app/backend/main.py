"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    main.py

Descripción:
    Punto de entrada de la aplicación.
"""

from services.importer import Importer
from services.synchronizer import Synchronizer


def main() -> None:
    """
    Ejecuta la importación y sincronización completa.
    """

    importer = Importer()

    pdf_files = importer.run()

    if not pdf_files:
        return

    synchronizer = Synchronizer()

    synchronizer.run(
        pdf_files,
    )


if __name__ == "__main__":
    main()
