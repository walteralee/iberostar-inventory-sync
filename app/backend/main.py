"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    main.py

Descripción:
    Punto de entrada de la aplicación.
"""

from services.synchronizer import Synchronizer


def main() -> None:
    """
    Ejecuta la sincronización completa.
    """

    synchronizer = Synchronizer()
    synchronizer.run()


if __name__ == "__main__":
    main()